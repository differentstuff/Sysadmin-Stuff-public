function Update-GitRelationship {
    <#
    .SYNOPSIS
        Resolves common Git repository issues by removing desktop.ini files and resetting Git relationships.
    
    .DESCRIPTION
        This function helps fix Git repositories that have been corrupted by desktop.ini files or other issues.
        It can perform three main operations:
        1. Delete all desktop.ini files in the repository
        2. Reset the Git relationship with the remote repository
        3. Reset Git credentials with option to re-enter new credentials
        
    .PARAMETER DeleteIniOnly
        When specified, only removes desktop.ini files without resetting Git relationships.
        
    .PARAMETER FullGitReset
        When specified, removes desktop.ini files and resets the Git relationship to match origin/main.
        
    .PARAMETER ResetGitCreds
        When specified, clears Git credentials including username, email, and stored credentials.
        Prompts user to enter new credentials after clearing.
        
    .PARAMETER Confirm
        When specified, bypasses confirmation prompts for operations.
        
    .PARAMETER Path
        Specifies the path to the Git repository. Defaults to the current directory.
        
    .EXAMPLE
        Update-GitRelationship -DeleteIniOnly -Path C:\Projects\MyRepo
        # Removes all desktop.ini files in the specified repository
        
    .EXAMPLE
        Update-GitRelationship -FullGitReset
        # Removes desktop.ini files and resets Git relationship in the current directory
        
    .EXAMPLE
        Update-GitRelationship -ResetGitCreds -Confirm
        # Resets Git credentials without confirmation prompts and offers to set new credentials
        
    .NOTES
        Author: Security Researcher
        Version: 1.2
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, HelpMessage = "Only remove desktop.ini files without resetting Git")]
        [switch]$DeleteIniOnly,
    
        [Parameter(Mandatory = $false, HelpMessage = "Remove desktop.ini files and reset Git relationship")]
        [switch]$FullGitReset,
        
        [Parameter(Mandatory = $false, HelpMessage = "Reset Git credentials (username, email, and stored credentials)")]
        [switch]$ResetGitCreds,
    
        [Parameter(Mandatory = $false, HelpMessage = "Skip confirmation prompts")]
        [switch]$Confirm,

        [Parameter(Mandatory = $false, HelpMessage = "Path to the Git repository")]
        [ValidateScript({
            if (-not (Test-Path $_)) {
                throw "Path '$_' does not exist."
            }
            return $true
        })]
        [string]$Path = "."
    )

    function Remove-DesktopIniFiles {
        param(
            [string]$WorkingPath
        )
    
        try {
            $iniFiles = Get-ChildItem -Path $WorkingPath -Filter desktop.ini -Recurse -Force -File -ErrorAction Stop
            $fileCount = ($iniFiles | Measure-Object).Count
        
            Write-Host "Found $fileCount desktop.ini files to delete in '$WorkingPath'" -ForegroundColor Cyan
        
            if ($fileCount -eq 0) {
                Write-Host "No desktop.ini files found." -ForegroundColor Yellow
                return $true
            }
        
            if ($Confirm -or (Read-Host "Do you want to proceed with deleting these files? (Y/N)").ToUpper() -eq "Y") {
                $iniFiles | ForEach-Object {
                    try {
                        $filePath = $_.FullName
                        Remove-Item -Path $filePath -Force -ErrorAction Stop
                        Write-Verbose "Deleted: $filePath"
                    }
                    catch {
                        Write-Warning "Failed to delete file: $filePath. Error: $_"
                    }
                }
                Write-Host "Successfully deleted $fileCount desktop.ini files" -ForegroundColor Green
                return $true
            }
        
            Write-Host "File deletion cancelled by user." -ForegroundColor Yellow
            return $false
        }
        catch {
            Write-Error "Error while searching for desktop.ini files: $_"
            return $false
        }
    }

    function Reset-GitRelationship {
        param(
            [string]$WorkingPath
        )

        # Change to the specified directory
        try {
            Push-Location $WorkingPath -ErrorAction Stop
        
            if ($Confirm -or (Read-Host "This will reset your git relationship and fetch latest from origin/main. Continue? (Y/N)").ToUpper() -eq "Y") {
                try {
                    # Verify we're in a git repository
                    if (-not (Test-Path ".git" -ErrorAction Stop)) {
                        Write-Host "Error: '$WorkingPath' is not a git repository." -ForegroundColor Red
                        Pop-Location
                        return $false
                    }

                    Write-Host "Removing the problematic reference..." -ForegroundColor Cyan
                    $result = git update-ref -d refs/desktop.ini 2>&1
                    if ($LASTEXITCODE -ne 0) {
                        Write-Warning "Warning when removing reference: $result"
                    }
                
                    Write-Host "Fetching the latest state..." -ForegroundColor Cyan
                    $result = git fetch origin 2>&1
                    if ($LASTEXITCODE -ne 0) {
                        Write-Error "Failed to fetch from origin: $result"
                        Pop-Location
                        return $false
                    }
                
                    Write-Host "Resetting to match remote main branch..." -ForegroundColor Cyan
                    $result = git reset --hard origin/main 2>&1
                    if ($LASTEXITCODE -ne 0) {
                        Write-Error "Failed to reset to origin/main: $result"
                        Pop-Location
                        return $false
                    }
                
                    Write-Host "Git relationship successfully reset and updated" -ForegroundColor Green
                    Pop-Location
                    return $true
                }
                catch {
                    Write-Error "An error occurred during git operations: $_"
                    Pop-Location
                    return $false
                }
            }
        
            Pop-Location
            Write-Host "Git reset cancelled by user." -ForegroundColor Yellow
            return $false
        }
        catch {
            Write-Error "Failed to access directory '$WorkingPath': $_"
            return $false
        }
    }

    function Reset-GitCredentials {
        try {
            Write-Host "Resetting Git credentials..." -ForegroundColor Cyan
            
            # Store current credentials before removing them (for informational purposes)
            $currentUsername = git config --global user.name 2>$null
            $currentEmail = git config --global user.email 2>$null
            
            # Unset global username
            $result = git config --global --unset user.name 2>&1
            if ($LASTEXITCODE -ne 0 -and $result -notmatch "No such key") {
                Write-Warning "Warning when unsetting user.name: $result"
            }
            
            # Unset global email
            $result = git config --global --unset user.email 2>&1
            if ($LASTEXITCODE -ne 0 -and $result -notmatch "No such key") {
                Write-Warning "Warning when unsetting user.email: $result"
            }
            
            # Check if git-credential-manager is available
            $gcmAvailable = $null -ne (Get-Command "git-credential-manager" -ErrorAction SilentlyContinue)
            if ($gcmAvailable) {
                $result = git credential-manager erase 2>&1
                if ($LASTEXITCODE -ne 0) {
                    Write-Warning "Warning when erasing credentials: $result"
                }
            }
            else {
                Write-Warning "Git Credential Manager not found. Credentials may not be fully cleared."
            }
            
            Write-Host "Git credentials have been reset successfully" -ForegroundColor Green
            
            # Ask user if they want to enter new credentials
            $setupNewCreds = $Confirm
            if (-not $Confirm) {
                $response = Read-Host "Would you like to set up new Git credentials? (Y/N)"
                $setupNewCreds = $response.ToUpper() -eq "Y"
            }
            
            if ($setupNewCreds) {
                Write-Host "`nSetting up new Git credentials" -ForegroundColor Cyan
                
                # Show previous values as reference
                if (-not [string]::IsNullOrWhiteSpace($currentUsername)) {
                    Write-Host "Previous username: $currentUsername" -ForegroundColor Yellow
                }
                if (-not [string]::IsNullOrWhiteSpace($currentEmail)) {
                    Write-Host "Previous email: $currentEmail" -ForegroundColor Yellow
                }
                
                # Get new username
                $newUsername = Read-Host "Enter your Git username"
                if (-not [string]::IsNullOrWhiteSpace($newUsername)) {
                    $result = git config --global user.name "$newUsername" 2>&1
                    if ($LASTEXITCODE -ne 0) {
                        Write-Error "Failed to set username: $result"
                    }
                    else {
                        Write-Host "Username set successfully" -ForegroundColor Green
                    }
                }
                else {
                    Write-Host "Username not provided, skipping..." -ForegroundColor Yellow
                }
                
                # Get new email
                $newEmail = Read-Host "Enter your Git email"
                if (-not [string]::IsNullOrWhiteSpace($newEmail)) {
                    $result = git config --global user.email "$newEmail" 2>&1
                    if ($LASTEXITCODE -ne 0) {
                        Write-Error "Failed to set email: $result"
                    }
                    else {
                        Write-Host "Email set successfully" -ForegroundColor Green
                    }
                }
                else {
                    Write-Host "Email not provided, skipping..." -ForegroundColor Yellow
                }
                
                # Inform about credential manager
                if ($gcmAvailable) {
                    Write-Host "`nGit Credential Manager is available on your system." -ForegroundColor Cyan
                    Write-Host "Your credentials will be requested the next time you perform a Git operation requiring authentication." -ForegroundColor Cyan
                }
                else {
                    Write-Host "`nGit Credential Manager is not available on your system." -ForegroundColor Yellow
                    Write-Host "You may need to enter your credentials manually for Git operations requiring authentication." -ForegroundColor Yellow
                }
                
                Write-Host "`nNew Git credentials have been configured" -ForegroundColor Green
            }
            else {
                Write-Host "Skipped setting up new Git credentials" -ForegroundColor Yellow
            }
            
            return $true
        }
        catch {
            Write-Error "Failed to reset Git credentials: $_"
            return $false
        }
    }

    # Main execution logic
    $operationSelected = $DeleteIniOnly -or $FullGitReset -or $ResetGitCreds
    
    if (-not $operationSelected) {
        Write-Host "Please specify at least one operation: -DeleteIniOnly, -FullGitReset, or -ResetGitCreds" -ForegroundColor Yellow
        Write-Host "Use Get-Help Update-GitRelationship for more information." -ForegroundColor Cyan
        return
    }

    try {
        # Resolve to absolute path
        $absolutePath = Resolve-Path $Path -ErrorAction Stop
        Write-Host "Working with path: $absolutePath" -ForegroundColor Cyan

        $results = @()

        if ($DeleteIniOnly) {
            $iniResult = Remove-DesktopIniFiles -WorkingPath $absolutePath
            $results += [PSCustomObject]@{
                Operation = "Remove desktop.ini files"
                Success = $iniResult
            }
        }

        if ($FullGitReset) {
            $iniResult = Remove-DesktopIniFiles -WorkingPath $absolutePath
            $results += [PSCustomObject]@{
                Operation = "Remove desktop.ini files"
                Success = $iniResult
            }
            
            if ($iniResult) {
                $gitResult = Reset-GitRelationship -WorkingPath $absolutePath
                $results += [PSCustomObject]@{
                    Operation = "Reset Git relationship"
                    Success = $gitResult
                }
            }
        }

        if ($ResetGitCreds) {
            $credsResult = Reset-GitCredentials
            $results += [PSCustomObject]@{
                Operation = "Reset Git credentials"
                Success = $credsResult
            }
        }

        # Display summary
        Write-Host "`nOperation Summary:" -ForegroundColor Cyan
        $results | Format-Table -AutoSize
        
        # Check if all operations succeeded
        $allSucceeded = $results | Where-Object { -not $_.Success } | Measure-Object | Select-Object -ExpandProperty Count -eq 0
        
        if ($allSucceeded) {
            Write-Host "All operations completed successfully." -ForegroundColor Green
        }
        else {
            Write-Host "Some operations failed. Please review the summary above." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Error "An unexpected error occurred: $_"
    }
}