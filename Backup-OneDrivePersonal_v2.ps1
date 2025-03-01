function Backup-OneDrivePersonal{
    <#
    .SYNOPSIS
        OneDrive Personal Backup and Management Tool
    .DESCRIPTION
        This script provides a user-friendly interface to manage your OneDrive Personal account.
        It allows you to view OneDrive statistics, back up content, and verify existing backups.
        Supports both PowerShell 5.1 (sequential processing) and PowerShell 7+ (parallel processing).
    .PARAMETER UpdateModules
        If specified, updates the Microsoft Graph PowerShell modules.
    .EXAMPLE
        .\Backup-OneDrivePersonal.ps1
    .EXAMPLE
        .\Backup-OneDrivePersonal.ps1 -UpdateModules
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [switch]$UpdateModules
    )

#region Variables

    # Global Level variables
    $global:StartTime = Get-Date
    $global:ProcessedFilesLock = [System.Object]::new()
    $global:SkippedFilesLock = [System.Object]::new()
    $global:ErrorCountLock = [System.Object]::new()
    $global:Statistics = [System.Collections.Concurrent.ConcurrentDictionary[string,int]]::new()

    # Script-level variables
    $Script:OverwriteAll = $false
    $Script:TotalFiles = 0
    $Script:IsConnected = $false
    $Script:ExcludedFolders = @()
    $script:MaxDepth = 8
    $Script:ThrottleLimit = 10
    $Script:QueueThrottleLimit = 10
    $script:WaitCPU = $false  # whether to wait
    $script:WaitCPUTime = 100  # Value in ms
    $Script:RequestTimestamps = [System.Collections.ArrayList]::new()
    $Script:MaxRequestsPerMinute = 600  # Adjust based on Microsoft Graph limits
    $Script:ErrorItems = [System.Collections.Concurrent.ConcurrentBag[object]]::new()

    # Tracking
    $global:syncHash = [hashtable]::Synchronized(@{
        ProcessedFiles = 0
        SkippedFiles = 0
        ErrorCount = 0
    })

#endregion Variables


#region UI Functions

    # Function to display welcome message and menu
    function Show-WelcomeMenu {
        Clear-Host
        Write-Host "=======================================================" -ForegroundColor Cyan
        Write-Host "             ONEDRIVE PERSONAL MANAGEMENT TOOL         " -ForegroundColor Cyan
        Write-Host "=======================================================" -ForegroundColor Cyan
        Write-Host "*made by differentstuff*"
        Write-Host ""
        Write-Host "Welcome to the OneDrive Personal Management Tool!"
        Write-Host "This tool helps you manage, back up, and verify your OneDrive content."
        Write-Host ""
        Write-Host "Please select an option:" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  [1] View OneDrive Statistics" -ForegroundColor Green
        Write-Host "  [2] Make a Full Backup" -ForegroundColor Green
        Write-Host "  [3] Verify Existing Backup" -ForegroundColor Green
        Write-Host "  [4] Manage Folder Exclusions" -ForegroundColor Green 
        Write-Host "  [5] Change OneDrive Account" -ForegroundColor Green
        Write-Host "  [0] Exit" -ForegroundColor Green
        Write-Host ""
    
        do {
            $choice = Read-Host "Enter your choice (0-5)"
        } while ($choice -lt 0 -or $choice -gt 5)
    
        switch ($choice) {
            1 { Show-OneDriveStatistics }
            2 { Start-FullBackup }
            3 { Start-BackupVerification }
            4 { Set-FolderExclusions }
            5 { Select-OneDriveAccount }
            0 { 
                Write-Host "Thank you for using the OneDrive Personal Management Tool!" -ForegroundColor Cyan
                return # Return from function instead of exiting the script
            }
        }
    }

# function for spinner-based progress
function Show-SpinnerProgress {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Activity,
        
        [Parameter(Mandatory = $false)]
        [string]$Status = "",
        
        [Parameter(Mandatory = $false)]
        [int]$PercentComplete = -1,
        
        [Parameter(Mandatory = $false)]
        [switch]$NoNewLine
    )
    
    # Static spinner characters
    $spinChars = @('|', '/', '-', '\')
    
    # Use a script-scoped variable to track the spinner position
    if (-not $Script:SpinnerIndex) {
        $Script:SpinnerIndex = 0
    }
    
    $spinChar = $spinChars[$Script:SpinnerIndex % 4]
    $Script:SpinnerIndex++
    
    # Build the progress message
    $message = "$Activity"
    if ($Status) {
        $message += " - $Status"
    }
    if ($PercentComplete -ge 0) {
        $message += " [$PercentComplete%]"
    }
    $message += " $spinChar"
    
    # Display the progress
    if ($NoNewLine) {
        Write-Host "`r$message" -NoNewline
    } else {
        Write-Host "`r$message"
    }
}


#endregion UI Functions


#region Local Functions

    # Function to format file size in a readable way
    function Format-FileSize {
        param (
            [Parameter(Mandatory = $true)]
            [long]$SizeInBytes
        )
    
        if ($SizeInBytes -ge 1TB) {
            return "{0:N2} TB" -f ($SizeInBytes / 1TB)
        }
        elseif ($SizeInBytes -ge 1GB) {
            return "{0:N2} GB" -f ($SizeInBytes / 1GB)
        }
        elseif ($SizeInBytes -ge 1MB) {
            return "{0:N2} MB" -f ($SizeInBytes / 1MB)
        }
        elseif ($SizeInBytes -ge 1KB) {
            return "{0:N2} KB" -f ($SizeInBytes / 1KB)
        }
        else {
            return "$SizeInBytes Bytes"
        }
    }

    # Function to initialize Graph modules
    function Initialize-GraphModules {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $false)]
            [switch]$UpdateModules
        )
    
        try {
            $requiredModules = @(
                "Microsoft.Graph",
                "Microsoft.Graph.Authentication",
                "Microsoft.Graph.Files"
            )
        
            foreach ($module in $requiredModules) {
                if (Get-Module -ListAvailable -Name $module) {
                    Write-Verbose "Module $module is installed"
                    if ($UpdateModules) {
                        Write-Host "Updating module $module..."
                        Update-Module $module -ErrorAction Stop
                    }
                }
                else {
                    Write-Host "Installing module $module..."
                    Install-Module $module -Scope CurrentUser -Force -ErrorAction Stop
                }
            }
            return $true
        }
        catch {
            Write-Error -Exception ([System.Management.Automation.PSInvalidOperationException]::new(
                "Failed to initialize Graph modules: $_",
                $_.Exception
            ))
            return $false
        }
    }

    # Function to connect to OneDrive
    function Connect-ToOneDrive {
        [CmdletBinding()]
        param(
            # Define required scopes for OneDrive access
            [Parameter(Mandatory=$false)]
            [System.Collections.Generic.List[string]]
            $Scopes = @(
                "Files.Read",
                "Files.Read.All", 
                "Sites.Read.All",
                "User.Read"
            ),
            
            [Parameter(Mandatory=$false)]
            [string]
            $TenantId = "consumers",
            
            [Parameter(Mandatory=$false)]
            [switch]
            $Silent,
            
            [Parameter(Mandatory=$false)]
            [System.Management.Automation.PSCredential]
            $Credential
        )
        
        try {
            # Check if already connected with required scopes
            $currentContext = Get-MgContext -ErrorAction SilentlyContinue
            if ($currentContext) {
                $hasAllScopes = $true
                foreach ($scope in $Scopes) {
                    if ($currentContext.Scopes -notcontains $scope) {
                        $hasAllScopes = $false
                        break
                    }
                }
                
                if ($hasAllScopes) {
                    # We're already connected, get user info
                    try {
                        $userInfo = Get-MgUser -UserId "me" -ErrorAction Stop
                        $userEmail = $userInfo.UserPrincipalName
                        $userName = $userInfo.DisplayName
                        $userID = $userInfo.Id
                        
                        if (-not $Silent) {
                            Write-Host "Already connected to Microsoft Graph as $userName ($userEmail) with required scopes" -ForegroundColor Green
                        }
                        
                        # Store user info in script variables for later use
                        $Script:UserEmail = $userEmail
                        $Script:UserName = $userName
                        $Script:userID = $userID
                        $Script:IsConnected = $true
                        return $true
                    }
                    catch {
                        # If we can't get user info, disconnect and try again
                        if (-not $Silent) {
                            Write-Host "Connected but unable to retrieve user information. Reconnecting..." -ForegroundColor Yellow
                        }
                        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                    }
                }
                else {
                    # Disconnect if we don't have all required scopes
                    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                }
            }
            
            # Connect parameters
            $connectParams = @{
                Scopes = $Scopes
                TenantId = $TenantId
                ErrorAction = "Stop"
            }
            
            # Add credential if provided
            if ($Credential) {
                $connectParams.Add("Credential", $Credential)
            }
            
            # Try to connect
            if (-not $Silent) {
                Write-Host "Connecting to Microsoft Graph..."
            }
            
            try{
                Connect-MgGraph @connectParams | Out-Null
            }
            catch {
                # Properly disconnect
                Disconnect-MgGraph -ErrorAction -ProgressAction | Out-Null

                # Clear any cached tokens
                [Microsoft.Identity.Client.PublicClientApplication]::ClearTokenCache()

                # Try to connect again
                Connect-MgGraph @connectParams | Out-Null
            }
            
            # Verify connection
            $context = Get-MgContext
            if (-not $context) {
                throw "Failed to establish Microsoft Graph connection"
            }
            
            # Get user information
            try {
                $userInfo = Get-MgUser -UserId "me" -ErrorAction Stop
                $userEmail = $userInfo.UserPrincipalName
                $userName = $userInfo.DisplayName
                
                # Store user info in script variables for later use
                $Script:UserEmail = $userEmail
                $Script:UserName = $userName
                
                if (-not $Silent) {
                    Write-Host "Connected to Microsoft Graph as $userName ($userEmail)" -ForegroundColor Green
                }
            }
            catch {
                # If we can't get user info but are connected, still proceed
                if (-not $Silent) {
                    Write-Host "Connected to Microsoft Graph, but unable to retrieve user details: $_" -ForegroundColor Yellow
                    Write-Host "Connection established with ClientId: $($context.ClientId)" -ForegroundColor Green
                }
            }
            
            # Verify we can access OneDrive
            try {
                $drive = Get-MgDrive -ErrorAction Stop
                if (-not $drive) {
                    throw "Connected to Graph but unable to access OneDrive"
                }
                
                # Store drive info in script variable
                $Script:DriveInfo = $drive
                
                if (-not $Silent) {
                    Write-Host "OneDrive access verified: $($drive.DriveType)" -ForegroundColor Green
                }
            }
            catch {
                throw "Connected to Graph but unable to access OneDrive: $_"
            }
            
            $Script:IsConnected = $true
            return $true
        }
        catch {
            if (-not $Silent) {
                Write-Host "Authentication failed: $_" -ForegroundColor Red
            }
            $Script:IsConnected = $false
            return $false
        }
    }

    # Function to manage folder exclusions
    function Set-FolderExclusions {
        Clear-Host
        Write-Host "=======================================================" -ForegroundColor Cyan
        Write-Host "               MANAGE FOLDER EXCLUSIONS                " -ForegroundColor Cyan
        Write-Host "=======================================================" -ForegroundColor Cyan
        Write-Host ""
        
        Write-Host "Current excluded folders:" -ForegroundColor Yellow
        if ($Script:ExcludedFolders.Count -eq 0) {
            Write-Host "  No folders are currently excluded." -ForegroundColor Gray
        } else {
            for ($i = 0; $i -lt $Script:ExcludedFolders.Count; $i++) {
                Write-Host "  [$i] $($Script:ExcludedFolders[$i])" -ForegroundColor White
            }
        }
        
        Write-Host ""
        Write-Host "Options:" -ForegroundColor Yellow
        Write-Host "  [A] Add a folder to exclusion list" -ForegroundColor Green
        Write-Host "  [R] Remove a folder from exclusion list" -ForegroundColor Green
        Write-Host "  [C] Clear all exclusions" -ForegroundColor Green
        Write-Host "  [M] Return to main menu" -ForegroundColor Green
        Write-Host ""
        
        $choice = Read-Host "Enter your choice"
        
        switch ($choice.ToUpper()) {
            "A" {
                $newExclusion = Read-Host "Enter folder path to exclude (e.g., 'Documents' or 'Pictures/Vacation')"
                if ($newExclusion) {
                    # Normalize path format - ensure it doesn't start with / but can contain /
                    $newExclusion = $newExclusion.Trim().TrimStart("/")
                    if ($Script:ExcludedFolders -notcontains $newExclusion) {
                        $Script:ExcludedFolders += $newExclusion
                        Write-Host "Added '$newExclusion' to exclusions." -ForegroundColor Green
                    } else {
                        Write-Host "'$newExclusion' is already in the exclusion list." -ForegroundColor Yellow
                    }
                }
                Set-FolderExclusions
            }
            "R" {
                if ($Script:ExcludedFolders.Count -eq 0) {
                    Write-Host "No exclusions to remove." -ForegroundColor Yellow
                    Start-Sleep -Seconds 2
                } else {
                    $indexToRemove = Read-Host "Enter the index number to remove [0-$($Script:ExcludedFolders.Count - 1)]"
                    if ($indexToRemove -match '^\d+$' -and [int]$indexToRemove -ge 0 -and [int]$indexToRemove -lt $Script:ExcludedFolders.Count) {
                        $removed = $Script:ExcludedFolders[$indexToRemove]
                        $Script:ExcludedFolders = $Script:ExcludedFolders | Where-Object { $_ -ne $removed }
                        Write-Host "Removed '$removed' from exclusions." -ForegroundColor Green
                    } else {
                        Write-Host "Invalid index." -ForegroundColor Red
                    }
                }
                Set-FolderExclusions
            }
            "C" {
                $Script:ExcludedFolders = @()
                Write-Host "All exclusions have been cleared." -ForegroundColor Green
                Start-Sleep -Seconds 2
                Set-FolderExclusions
            }
            "M" {
                Show-WelcomeMenu
            }
            default {
                Write-Host "Invalid choice. Please try again." -ForegroundColor Red
                Start-Sleep -Seconds 2
                Set-FolderExclusions
            }
        }
    }


    # Function to check if a path should be excluded
    function Test-ShouldExcludePath {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$Path
        )
        
        # Normalize path - remove leading and trailing slashes
        $normalizedPath = $Path.Trim('/').Replace('\', '/')
        
        foreach ($exclusion in $Script:ExcludedFolders) {
            $exclusionPattern = $exclusion.Trim('/').Replace('\', '/')
            
            # Check for exact match or if path starts with exclusion pattern followed by /
            if ($normalizedPath -eq $exclusionPattern -or 
                $normalizedPath.StartsWith("$exclusionPattern/")) {
                Write-Verbose "Path '$Path' matched exclusion pattern '$exclusion'"
                return $true
            }
        }
        
        return $false
    }

    # Function to verify OneDrive backup
    function Test-OneDriveBackup {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$BackupFolder,
        
            [Parameter(Mandatory = $false)]
            [switch]$ExcludeSharedItems
        )
    
        # Display account information
        Write-Host ""
        if ($Script:UserName -and $Script:UserEmail) {
            Write-Host "Verifying OneDrive backup for account: $($Script:UserName) ($($Script:UserEmail))" -ForegroundColor Yellow
        } else {
            Write-Host "Verifying OneDrive backup..." -ForegroundColor Yellow
        }
        Write-Host "This may take several minutes for large backups." -ForegroundColor Yellow
        Write-Host ""
        
        # Show exclusions
        if ($Script:ExcludedFolders.Count -gt 0) {
            Write-Host "The following folders are excluded from verification:" -ForegroundColor Yellow
            foreach ($exclusion in $Script:ExcludedFolders) {
                Write-Host "  - $exclusion" -ForegroundColor Yellow
            }
            Write-Host ""
        }

        # Store verification results
        $results = @{
            Missing    = @()
            Modified   = @()
            Extra      = @()
            Identical  = 0
        }
    
        # Get OneDrive items
        $rootItems = Get-MgDriveRootChild -DriveId $($Script:DriveInfo.Id) -ErrorAction Stop
    
        # Create a mapping of OneDrive files
        $oneDriveFiles = @{}
        $verificationStartTime = Get-Date
    
        # Create a progress indicator
        $progressId = 2
        Show-SpinnerProgress -Activity "Verification" -Status "Building OneDrive file index..." -PercentComplete 0 -NoNewLine
    
        # Build the complete OneDrive file index
        $indexedItems = 0
        $oneDriveFiles = Build-OneDriveFileIndex -RootItems $rootItems -RootPath "\" -OneDriveFiles $oneDriveFiles -ProgressId $progressId -ref ([ref]$indexedItems) -IncludeSharedItems:$IncludeSharedItems
    
        Show-SpinnerProgress -Activity "Verification" -Status "Comparing with local backup..." -PercentComplete 50 -NoNewLine
    
        # Compare local files with OneDrive index
        $localFiles = Get-ChildItem $BackupFolder -Recurse -File
        $totalLocalFiles = $localFiles.Count
        $processedLocalFiles = 0
    
        foreach ($localFile in $localFiles) {
            $processedLocalFiles++
            $percentComplete = 50 + [math]::Round(($processedLocalFiles / $totalLocalFiles) * 50)
            Show-SpinnerProgress -Activity "Verification" -Status "Comparing file $processedLocalFiles of $totalLocalFiles" -PercentComplete $percentComplete -NoNewLine
        
            # Get relative path
            $relativePath = $localFile.FullName.Substring($BackupFolder.Length).Replace("\", "/")
            if (-not $relativePath.StartsWith("/")) {
                $relativePath = "/" + $relativePath
            }
        
            if ($oneDriveFiles.ContainsKey($relativePath)) {
                $oneDriveFile = $oneDriveFiles[$relativePath]
            
                # Compare sizes
                if ($localFile.Length -ne $oneDriveFile.Size) {
                    $sizeDiff = $oneDriveFile.Size - $localFile.Length
                    $formattedSizeDiff = if ($sizeDiff -gt 0) {
                        "+$(Format-FileSize $sizeDiff)"
                    } else {
                        "-$(Format-FileSize ([math]::Abs($sizeDiff)))"
                    }
                
                    $results.Modified += [PSCustomObject]@{
                        Path = $relativePath
                        LocalSize = Format-FileSize $localFile.Length
                        OneDriveSize = Format-FileSize $oneDriveFile.Size
                        SizeDifference = $formattedSizeDiff
                        LastModified = $oneDriveFile.LastModified
                        LocalLastModified = $localFile.LastWriteTime
                    }
                }
                else {
                    $results.Identical++
                }
            
                # Mark as processed
                $oneDriveFiles.Remove($relativePath)
            }
            else {
                $results.Extra += [PSCustomObject]@{
                    Path = $relativePath
                    Size = Format-FileSize $localFile.Length
                    LastModified = $localFile.LastWriteTime
                }
            }
        }
    
        # Any remaining files in OneDrive are missing from local backup
        foreach ($key in $oneDriveFiles.Keys) {
            $results.Missing += [PSCustomObject]@{
                Path = $key
                Size = Format-FileSize $oneDriveFiles[$key].Size
                LastModified = $oneDriveFiles[$key].LastModified
            }
        }
    
        # Complete progress indicator
        Write-Host "" # Clear the spinner line
    
        # Calculate verification time
        $verificationTime = (Get-Date) - $verificationStartTime
        $verificationTimeText = "{0:hh\:mm\:ss}" -f $verificationTime
    
        # Display verification results
        Clear-Host
        Write-Host "=======================================================" -ForegroundColor Cyan
        Write-Host "              BACKUP VERIFICATION RESULTS              " -ForegroundColor Cyan
        Write-Host "=======================================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Verification completed in $verificationTimeText" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "SUMMARY:" -ForegroundColor Green
        Write-Host "  Identical Files: $($results.Identical)"
        Write-Host "  Modified Files:  $($results.Modified.Count)"
        Write-Host "  Missing Files:   $($results.Missing.Count)"
        Write-Host "  Extra Files:     $($results.Extra.Count)"
        Write-Host ""
    
        if ($results.Modified.Count -gt 0) {
            Write-Host "MODIFIED FILES:" -ForegroundColor Yellow
            foreach ($file in $results.Modified) {
                Write-Host "  Path: $($file.Path)" -ForegroundColor White
                Write-Host "    Local Size:      $($file.LocalSize)"
                Write-Host "    OneDrive Size:   $($file.OneDriveSize)"
                Write-Host "    Size Difference: $($file.SizeDifference)"
                Write-Host "    OneDrive Modified: $($file.LastModified)"
                Write-Host "    Local Modified:    $($file.LocalLastModified)"
                Write-Host ""
            }
        }
    
        if ($results.Missing.Count -gt 0) {
            Write-Host "MISSING FILES (in OneDrive but not in backup):" -ForegroundColor Red
            foreach ($file in $results.Missing) {
                Write-Host "  Path: $($file.Path)" -ForegroundColor White
                Write-Host "    Size:         $($file.Size)"
                Write-Host "    Last Modified: $($file.LastModified)"
                Write-Host ""
            }
        }
    
        if ($results.Extra.Count -gt 0) {
            Write-Host "EXTRA FILES (in backup but not in OneDrive):" -ForegroundColor Magenta
            foreach ($file in $results.Extra) {
                Write-Host "  Path: $($file.Path)" -ForegroundColor White
                Write-Host "    Size:         $($file.Size)"
                Write-Host "    Last Modified: $($file.LastModified)"
                Write-Host ""
            }
        }
    
        if ($results.Modified.Count -eq 0 -and $results.Missing.Count -eq 0 -and $results.Extra.Count -eq 0) {
            Write-Host "Congratulations! Your backup is an exact match with OneDrive." -ForegroundColor Green
        }
    }


#endregion Local Functions


#region Onedrive Functions

    function Update-ThrottleLimit {
        [CmdletBinding()]
        param()
        
        try {
            $cpuLoad = (Get-Counter -Counter "\Processor(_Total)\% Processor Time" -ErrorAction SilentlyContinue).CounterSamples.CookedValue
            $memoryAvailable = (Get-Counter -Counter "\Memory\Available MBytes" -ErrorAction SilentlyContinue).CounterSamples.CookedValue
            
            # Default values if counters aren't available
            if ($null -eq $cpuLoad) { $cpuLoad = 50 }
            if ($null -eq $memoryAvailable) { $memoryAvailable = 4000 }
            
            # Adjust throttle limit based on system load
            if ($cpuLoad -gt 80 -or $memoryAvailable -lt 1000) {
                # High load - reduce throttle
                $newLimit = [Math]::Max(1, $Script:ThrottleLimit / 2)
                if ($newLimit -ne $Script:ThrottleLimit) {
                    $Script:ThrottleLimit = [int]$newLimit
                    Write-Verbose "Reducing throttle limit to $Script:ThrottleLimit due to system load (CPU: $cpuLoad%, Mem: $memoryAvailable MB)"
                }
            }
            elseif ($cpuLoad -lt 40 -and $memoryAvailable -gt 4000) {
                # Low load - increase throttle up to original max
                $newLimit = [Math]::Min($Script:QueueThrottleLimit, $Script:ThrottleLimit + 2)
                if ($newLimit -ne $Script:ThrottleLimit) {
                    $Script:ThrottleLimit = [int]$newLimit
                    Write-Verbose "Increasing throttle limit to $Script:ThrottleLimit (CPU: $cpuLoad%, Mem: $memoryAvailable MB)"
                }
            }
        }
        catch {
            Write-Verbose "Error updating throttle limit: $_"
        }
    }

    # Function to select/change OneDrive account
    function Select-OneDriveAccount {
        Clear-Host
        Write-Host "=======================================================" -ForegroundColor Cyan
        Write-Host "    SELECT ONEDRIVE ACCOUNT    " -ForegroundColor Cyan
        Write-Host "=======================================================" -ForegroundColor Cyan
        Write-Host ""
        
        # First disconnect from current account if connected
        if ($Script:IsConnected) {
            Write-Host "Disconnecting from current account: $($Script:UserEmail)..." -ForegroundColor Yellow
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            $Script:IsConnected = $false
            $Script:UserEmail = $null
            $Script:UserName = $null
            $Script:DriveInfo = $null
            Write-Host "Disconnected successfully." -ForegroundColor Green
        }
        
        # Connect to new account
        Write-Host "Connecting to a different OneDrive account..." -ForegroundColor Yellow
        if (Connect-ToOneDrive) {
            Write-Host "Successfully connected to new account: $($Script:UserEmail)" -ForegroundColor Green
            
            # Display drive information if available
            Show-DriveStorageInformation

        } else {
            Write-Host "Failed to connect to a new account." -ForegroundColor Red
        }
        
        Write-Host ""
        Write-Host "Press any key to return to the main menu..."
        [void][System.Console]::ReadKey($true)
        Show-WelcomeMenu
    }

    # Function to get all files inside a path from Onedrive
    function Get-OneDriveItems {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $false)]
            [string]$FolderId = "root",
            
            [Parameter(Mandatory = $false)]
            [string]$Path = ""
        )
        
        try {
            $uri = if ($FolderId -eq "root") {
                "https://graph.microsoft.com/v1.0/me/drive/root/children"
            } else {
                "https://graph.microsoft.com/v1.0/me/drive/items/$FolderId/children"
            }
            $response = Invoke-MgGraphRequest -Uri $uri -Method GET
            $items = $response.value
            
            # Process items to clearly identify type
            $processedItems = @()
            
            foreach ($item in $items) {
                # Initialize variables
                $type = ""
                $isShared = $false
                $isEmpty = $false

                # Check for file (has @microsoft.graph.downloadUrl or file property)
                if ($item.file) {
                    $type = "File"
                    if($item.size -eq 0){
                        $isEmpty = $true
                    }
                }
                # Check for folder (has folder property)
                elseif ($item.folder.childCount) {
                    $type = "Folder"
                    if ($item.folder.childCount -eq 0){
                        $isEmpty = $true
                    }
                }
                # Check for remoteItem (shared items)
                elseif ($item.remoteItem) {
                    $type = "SharedItem"
                    $isShared = $true
                    if($item.remoteItem.size -eq 0){
                        $isEmpty = $true                        
                    }
                }
                
                # Create custom object with properties
                $processedItem = [PSCustomObject]@{
                    Id = $item.id
                    Name = $item.name
                    Path = if ($Path) { Join-Path -Path $Path -ChildPath $item.name } else { $item.name }
                    Size = $item.size
                    LastModifiedDateTime = $item.lastModifiedDateTime
                    Type = $type
                    IsShared = $isShared
                    IsEmpty = $isEmpty
                    DriveItem = $item
                }
                
                $processedItems += $processedItem
            }
            
            return $processedItems
        }
        catch {
            Write-Error "Failed to get OneDrive items: $_"
            return @()
        }
    }

    # Function to count total files in OneDrive with exclusion support
    function Get-OneDriveTotalItems {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            [switch]$IgnoreExclusions
        )
        
        $count = 0
        $progressId = 10
        
        # Get root items
        Show-SpinnerProgress -Activity "Counting files" -Status "Getting root items..." -PercentComplete 0 -NoNewLine
        $rootItems = Get-MgDriveRootChild -DriveId $($Script:DriveInfo.Id) -ErrorAction Stop
        
        # Count files recursively
        $count = Get-OneDriveFilesCount -RootItems $rootItems -RootPath "\" -ProgressId $progressId -IgnoreExclusions:$IgnoreExclusions
        
        Write-Host "" # Clear the spinner line
        return $count
    }
    
    function Get-OneDriveFilesCount {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [object[]]$RootItems,
            
            [Parameter(Mandatory = $true)]
            [string]$RootPath,
            
            [Parameter(Mandatory = $true)]
            [int]$ProgressId,
            
            [Parameter(Mandatory = $false)]
            [switch]$IgnoreExclusions
        )
        
        $fileCount = 0
        
        foreach ($item in $RootItems) {
            $path = Join-Path -Path $RootPath -ChildPath $item.Name -PathType NoteProperty
            $path = $path.Replace("\", "/")
            
            # Skip excluded paths
            $relativePath = $path.TrimStart('/')
            if (-not $IgnoreExclusions -and (Test-ShouldExcludePath -Path $relativePath)) {
                Write-Verbose "Excluding from count: $relativePath"
                continue
            }
            
            if ($item.Folder) {
                # Process folder recursively
                $children = Get-MgDriveItemChild -DriveId $($Script:DriveInfo.Id) -DriveItemId $item.Id -ErrorAction Stop
                $fileCount += Count-OneDriveFiles -RootItems $children -RootPath $path -ProgressId $ProgressId -IgnoreExclusions:$IgnoreExclusions
            }
            else {
                # It's a file, increment count
                $fileCount++
                
                if ($fileCount % 100 -eq 0) {
                    Show-SpinnerProgress -Activity "Counting files" -Status "Found $fileCount files so far..." -PercentComplete -1 -NoNewLine
                }
            }
        }
        
        return $fileCount
    }

    # Function to get folder statistics with depth limitation
    function Get-FolderStatistics {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$DriveItemId,
            
            [Parameter(Mandatory = $true)]
            [string]$FolderName,
            
            [Parameter(Mandatory = $false)]
            [string]$ParentPath = "/",
            
            [Parameter(Mandatory = $false)]
            [int]$MaxDepth = 3,
            
            [Parameter(Mandatory = $false)]
            [int]$CurrentDepth = 0
        )

        $stats = @{
            TotalFiles     = 0
            TotalFolders   = 0
            TotalSize      = 0
            FileTypes      = @{}
            LargestFiles   = @()
        }

        $currentPath = if ($ParentPath -eq "/") { "/$FolderName" } else { "$ParentPath/$FolderName" }

        try {
            # Get items in the folder
            $children = Get-MgDriveItemChild -DriveId $Script:DriveInfo.Id -DriveItemId $DriveItemId -ErrorAction Stop

            # Process each item
            foreach ($item in $children) {
                if ($item.Folder) {
                    $stats.TotalFolders++
                    
                    # If we've reached max depth, just get folder size without recursing
                    if ($CurrentDepth -ge $MaxDepth) {
                        # Add folder size from metadata without recursing
                        $stats.TotalSize += $item.Size
                        
                        # Estimate files based on average file size if available
                        if ($item.Folder.ChildCount -gt 0) {
                            $stats.TotalFiles += $item.Folder.ChildCount
                        }
                    } else {
                        # Recurse with incremented depth
                        $folderData = Get-FolderStatistics -DriveItemId $item.Id -FolderName $item.Name `
                                    -ParentPath $currentPath -MaxDepth $MaxDepth -CurrentDepth ($CurrentDepth + 1)
                        
                        $stats.TotalFiles += $folderData.TotalFiles
                        $stats.TotalFolders += $folderData.TotalFolders
                        $stats.TotalSize += $folderData.TotalSize
                        
                        # Add file types
                        foreach ($type in $folderData.FileTypes.Keys) {
                            if ($stats.FileTypes.ContainsKey($type)) {
                                $stats.FileTypes[$type] += $folderData.FileTypes[$type]
                            } else {
                                $stats.FileTypes[$type] = $folderData.FileTypes[$type]
                            }
                        }

                        # Add to largest files list
                        $stats.LargestFiles += $folderData.LargestFiles
                    }
                } else {
                    $stats.TotalFiles++
                    $stats.TotalSize += $item.Size
                    
                    # Add file type
                    $extension = [System.IO.Path]::GetExtension($item.Name).ToLower()
                    if ($extension) {
                        if ($stats.FileTypes.ContainsKey($extension)) {
                            $stats.FileTypes[$extension] += 1
                        } else {
                            $stats.FileTypes[$extension] = 1
                        }
                    } else {
                        $noExtension = "(no extension)"
                        if ($stats.FileTypes.ContainsKey($noExtension)) {
                            $stats.FileTypes[$noExtension] += 1
                        } else {
                            $stats.FileTypes[$noExtension] = 1
                        }
                    }
                    
                    # Add to largest files list if size is available
                    if ($item.Size -gt 0) {
                        $stats.LargestFiles += [PSCustomObject]@{
                            Name = $item.Name
                            Path = $currentPath
                            Size = $item.Size
                        }
                    }
                }
            }

            # Limit largest files to top 20 to avoid memory issues
            if ($stats.LargestFiles.Count -gt 20) {
                $stats.LargestFiles = $stats.LargestFiles | Sort-Object -Property Size -Descending | Select-Object -First 20
            }
        } catch {
            Write-Warning "Failed to process folder $($FolderName): $_"
        }

        return $stats
    }

    # Show Drive Storage Information
    function Show-DriveStorageInformation{
        # Display drive information if available
        if ($Script:DriveInfo) {
            Write-Host "Drive Information:" -ForegroundColor Yellow
            Write-Host "  Drive Type: $($Script:DriveInfo.DriveType)"
            Write-Host "  Drive ID: $($Script:DriveInfo.Id)"
            
            # Display quota information if available
            if ($Script:DriveInfo.Quota) {
                $usedGB = [math]::Round($Script:DriveInfo.Quota.Used / 1GB, 2)
                $totalGB = [math]::Round($Script:DriveInfo.Quota.Total / 1GB, 2)
                $percentUsed = if ($Script:DriveInfo.Quota.Total -gt 0) {
                    [math]::Round(($Script:DriveInfo.Quota.Used / $Script:DriveInfo.Quota.Total) * 100, 1)
                } else {
                    0
                }
                
                Write-Host "  Storage Used: $usedGB GB of $totalGB GB ($percentUsed%)"
                Write-Host "  State: $($Script:DriveInfo.Quota.State)"
            }
        }
    }
    
    # Function to display OneDrive statistics
    function Show-OneDriveStatistics {
        Clear-Host
        Write-Host "=======================================================" -ForegroundColor Cyan
        Write-Host "                ONEDRIVE STATISTICS                    " -ForegroundColor Cyan
        Write-Host "=======================================================" -ForegroundColor Cyan
        Write-Host ""
    
        # Ensure we're connected
        if (-not $Script:IsConnected) {
            if (-not (Connect-ToOneDrive)) {
                Write-Host "Unable to connect to OneDrive. Returning to main menu..." -ForegroundColor Red
                Start-Sleep -Seconds 3
                Show-WelcomeMenu
                return
            }
        }
    
        Write-Host "Analyzing your OneDrive content..." -ForegroundColor Yellow
        Write-Host "This may take several minutes for large accounts." -ForegroundColor Yellow
        Write-Host ""
        
        # Collect statistics
        $stats = @{
            TotalFiles       = 0
            TotalFolders     = 0
            TotalSize        = 0
            FileTypes        = @{}
            LargestFiles     = @()
            LargestFolders   = @{}
        }
    
        # Create a progress indicator
        $progressId = 1
        Show-SpinnerProgress -Activity "Analyzing OneDrive" -Status "Starting analysis..." -PercentComplete 0 -NoNewLine
    
        # Get root items
        $rootItems = Get-MgDriveRootChild -DriveId $($Script:DriveInfo.Id) -ErrorAction Stop
        $totalRootItems = $rootItems.Count
        $processedRootItems = 0
    
        # Process each root item
        foreach ($item in $rootItems) {
            $processedRootItems++
            $percentComplete = [math]::Round(($processedRootItems / $totalRootItems) * 100)
        
            if ($item.Folder) {
                Show-SpinnerProgress -Activity "Analyzing OneDrive" -Status "Analyzing folder: $($item.Name)" -PercentComplete $percentComplete -NoNewLine
                $folderData = Get-FolderStatistics -DriveItemId $item.Id -FolderName $item.Name -MaxDepth $script:MaxDepth
            
                $stats.TotalFiles += $folderData.TotalFiles
                $stats.TotalFolders += $folderData.TotalFolders + 1
                $stats.TotalSize += $folderData.TotalSize
            
                # Add file types
                foreach ($type in $folderData.FileTypes.Keys) {
                    if ($stats.FileTypes.ContainsKey($type)) {
                        $stats.FileTypes[$type] += $folderData.FileTypes[$type]
                    }
                    else {
                        $stats.FileTypes[$type] = $folderData.FileTypes[$type]
                    }
                }
            
                # Add to largest folders
                $stats.LargestFolders[$item.Name] = $folderData.TotalSize
            
                # Add to largest files list
                $stats.LargestFiles += $folderData.LargestFiles
            }
            else {
                Show-SpinnerProgress -Activity "Analyzing OneDrive" -Status "Analyzing file: $($item.Name)" -PercentComplete $percentComplete -NoNewLine
                $stats.TotalFiles++
                $stats.TotalSize += $item.Size
            
                # Add file type
                $extension = [System.IO.Path]::GetExtension($item.Name).ToLower()
                if ($extension) {
                    if ($stats.FileTypes.ContainsKey($extension)) {
                        $stats.FileTypes[$extension] += 1
                    }
                    else {
                        $stats.FileTypes[$extension] = 1
                    }
                }
                else {
                    $noExtension = "(no extension)"
                    if ($stats.FileTypes.ContainsKey($noExtension)) {
                        $stats.FileTypes[$noExtension] += 1
                    }
                    else {
                        $stats.FileTypes[$noExtension] = 1
                    }
                }
            
                # Add to largest files list if size is available
                if ($item.Size -gt 0) {
                    $stats.LargestFiles += [PSCustomObject]@{
                        Name = $item.Name
                        Path = "/"
                        Size = $item.Size
                    }
                }
            }
        }
    
        # Complete progress indicator
        Write-Host "" # Clear the spinner line
    
        # Sort largest files
        $stats.LargestFiles = $stats.LargestFiles | Sort-Object -Property Size -Descending | Select-Object -First 10
    
        # Sort largest folders
        $stats.LargestFolders = $stats.LargestFolders.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 10
    
        # Sort file types by count
        $stats.FileTypes = $stats.FileTypes.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 15
    
        # Display statistics
        Clear-Host
        Write-Host "=======================================================" -ForegroundColor Cyan
        Write-Host "                ONEDRIVE STATISTICS                    " -ForegroundColor Cyan
        Write-Host "=======================================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "OneDrive Account: $($Script:UserName) ($($Script:UserEmail))" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "SUMMARY:" -ForegroundColor Green
        Write-Host "  Total Files:   $($stats.TotalFiles)"
        Write-Host "  Total Folders: $($stats.TotalFolders)"
        Write-Host "  Total Size:    $(Format-FileSize $stats.TotalSize)"
        Write-Host ""
    
        Write-Host "TOP 10 LARGEST FILES:" -ForegroundColor Green
        $index = 1
        foreach ($file in $stats.LargestFiles) {
            Write-Host "  $index. $($file.Name) ($(Format-FileSize $file.Size))"
            $index++
        }
        Write-Host ""
    
        Write-Host "TOP 10 LARGEST FOLDERS:" -ForegroundColor Green
        $index = 1
        foreach ($folder in $stats.LargestFolders) {
            Write-Host "  $index. $($folder.Name) ($(Format-FileSize $folder.Value))"
            $index++
        }
        Write-Host ""
    
        Write-Host "TOP 15 FILE TYPES:" -ForegroundColor Green
        $index = 1
        foreach ($type in $stats.FileTypes) {
            Write-Host "  $index. $($type.Name): $($type.Value) files"
            $index++
        }
        Write-Host ""
    
        Write-Host "Press any key to return to the main menu..."
        [void][System.Console]::ReadKey($true)
        Show-WelcomeMenu
    }

    # Function to build OneDrive file index
    function Build-OneDriveFileIndex {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [object[]]$RootItems,
        
            [Parameter(Mandatory = $true)]
            [string]$RootPath,
        
            [Parameter(Mandatory = $true)]
            [hashtable]$OneDriveFiles,
        
            [Parameter(Mandatory = $true)]
            [int]$ProgressId,
        
            [Parameter(Mandatory = $true)]
            [ref]$IndexedItems,
        
            [Parameter(Mandatory = $false)]
            [switch]$IncludeSharedItems
        )
    
        foreach ($item in $RootItems) {
            # Skip shared items by default
            if (-not $IncludeSharedItems -and $item.remoteItem) {
                Write-Verbose -Message "Excluding all shared objects"
                continue
            }

            $path = Join-Path -Path $RootPath -ChildPath $item.Name -PathType NoteProperty
            $path = $path.Replace("\", "/")
        
            # Skip excluded paths
            $relativePath = $path.TrimStart('/')
            if (Test-ShouldExcludePath -Path $relativePath) {
                Write-Verbose "Excluding from verification: $relativePath"
                continue
            }
            
            if ($item.Folder) {
                # Process folder recursively
                $children = Get-MgDriveItemChild -DriveId $($Script:DriveInfo.Id) -DriveItemId $item.Id -ErrorAction Stop
                $OneDriveFiles = Build-OneDriveFileIndex -RootItems $children -RootPath $path -OneDriveFiles $OneDriveFiles -ProgressId $ProgressId -ref $IndexedItems
            }
            else {
                # It's a file, add to index
                $IndexedItems.Value++
            
                if ($IndexedItems.Value % 100 -eq 0) {
                    Show-SpinnerProgress -Activity "Verification" -Status "Indexing OneDrive files: $($IndexedItems.Value)" -PercentComplete 25 -NoNewLine
                }
            
                $OneDriveFiles[$path] = @{
                    Size = $item.Size
                    LastModified = $item.LastModifiedDateTime
                }
            }
        }
    
        return $OneDriveFiles
    }

    function Save-OneDriveFileBatch {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [object[]]$FileItems,
            
            [Parameter(Mandatory = $true)]
            [string]$DestinationRootPath,
            
            [Parameter(Mandatory = $false)]
            [switch]$Overwrite,
            
            [Parameter(Mandatory = $false)]
            [switch]$OverwriteAll,
            
            [Parameter(Mandatory = $false)]
            [int]$BatchSize = 15,
            
            [Parameter(Mandatory = $false)]
            [int]$MaxConcurrentConnections = 5,
            
            [Parameter(Mandatory = $false)]
            [int]$LargeFileThresholdMB = 100,
            
            [Parameter(Mandatory = $false)]
            [int]$ChunkSizeMB = 20,
            
            [Parameter(Mandatory = $false)]
            [int]$MaxRetries = 5
        )
        
        # Process items in batches
        for ($i = 0; $i -lt $FileItems.Count; $i += $BatchSize) {
            $batchItems = $FileItems | Select-Object -Skip $i -First $BatchSize
            
            $driveId = $Script:DriveInfo.Id

            # Process batch with parallel execution
            $batchItems | ForEach-Object -Parallel {
                $fileItem = $_
                
                # Extract file details
                $fileId = $fileItem.Id
                $destinationPath = $fileItem.Path
                $fileSize = $fileItem.Size
                $fileName = $fileItem.Name
                $lastModified = $fileItem.LastModifiedDateTime
                
                # Ensure directory exists
                $directory = Split-Path -Path $destinationPath -Parent
                if (-not (Test-Path $directory)) {
                    New-Item -Path $directory -ItemType Directory -Force | Out-Null
                }
                
                # Check if we should overwrite the file
                $shouldOverwrite = $true
                if (Test-Path $destinationPath) {
                    if ($using:OverwriteAll -or $global:OverwriteAll) {
                        $shouldOverwrite = $true
                    }
                    elseif ($using:Overwrite) {
                        $existingFile = Get-Item $destinationPath
                        $existingLastWrite = $existingFile.LastWriteTime.ToUniversalTime()
                        
                        if ($lastModified -and $lastModified -gt $existingLastWrite) {
                            $shouldOverwrite = $true
                        } else {
                            $shouldOverwrite = $false
                        }
                    }
                    else {
                        $shouldOverwrite = $false
                    }
                    
                    if (-not $shouldOverwrite) {
                        # Update skipped files counter
                        $global:syncHash.SkippedFiles++
                        return
                    }
                }
                
                # Retry logic setup
                $attempt = 0
                $success = $false
                $delay = 1  # Initial delay in seconds
                
                while (-not $success -and $attempt -lt $using:MaxRetries) {
                    $attempt++
                    try {
                        # Use a temp file for download to avoid partial downloads
                        $tempFile = [System.IO.Path]::GetTempFileName()
                        
                        # Regular file download using Graph API
                        Get-MgDriveItemContent -DriveId $using:driveId -DriveItemId $fileId -OutFile $tempFile -ErrorAction Stop

                        # Move temp file to destination
                        Move-Item -Path $tempFile -Destination $destinationPath -Force -ErrorAction Stop
                        
                        # Update the file's last modified time to match OneDrive
                        if ($lastModified) {
                            Set-ItemProperty -Path $destinationPath -Name LastWriteTime -Value $lastModified
                        }
                        
                        # Update processed files counter
                        $global:syncHash.ProcessedFiles++
                    
                    }
                    catch {
                        # Clean up temp file if it exists
                        if (Test-Path $tempFile) {
                            Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
                        }
                        
                        # Handle rate limiting (status code 429)
                        if ($_.Exception.Response -and $_.Exception.Response.StatusCode -eq 429) {
                            # Get retry-after header if available
                            $retryAfter = 1
                            if ($_.Exception.Response.Headers["Retry-After"]) {
                                $retryAfter = [int]$_.Exception.Response.Headers["Retry-After"]
                            }
                            
                            # Use exponential backoff with jitter
                            $delay = [Math]::Max($retryAfter, $delay * 2)
                            $jitter = Get-Random -Minimum 0 -Maximum 1000
                            $totalDelay = $delay + ($jitter / 1000)
                            
                            Write-Verbose "Rate limited. Retrying after $totalDelay seconds. Attempt $attempt/$using:MaxRetries"
                            Start-Sleep -Seconds $totalDelay
                        }
                        else {
                            # Other errors - use regular exponential backoff
                            if ($attempt -lt $using:MaxRetries) {
                                $delay = [Math]::Min(60, $delay * 2)  # Cap at 60 seconds
                                $jitter = Get-Random -Minimum 0 -Maximum 1000
                                $totalDelay = $delay + ($jitter / 1000)
                                
                                Write-Verbose "Error: $_. Retrying after $totalDelay seconds. Attempt $attempt/$using:MaxRetries"
                                Start-Sleep -Seconds $totalDelay
                            }
                            else {
                                # Final failure after all retries
                                Write-Error "Failed to download $fileName after $using:MaxRetries attempts: $_"

                                # Update error counter
                                $global:syncHash.ErrorCount++

                            }
                        }
                    }
                }
            } -ThrottleLimit $MaxConcurrentConnections
        }
    }
    
    # Helper function to handle large file downloads with chunking and resume capability
    function Save-LargeFile {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$DriveItemId,
            
            [Parameter(Mandatory = $true)]
            [string]$DestinationPath,
            
            [Parameter(Mandatory = $true)]
            [long]$FileSize,
            
            [Parameter(Mandatory = $false)]
            [int]$ChunkSizeMB = 20,
            
            [Parameter(Mandatory = $false)]
            [int]$MaxRetries = 3
        )
        
        try {
            # Create checkpoint file path for resume info
            $checkpointPath = "$DestinationPath.checkpoint"
            $chunkSizeBytes = $ChunkSizeMB * 1MB
            
            # Check if we have a checkpoint file to resume download
            $startPosition = 0
            if (Test-Path $checkpointPath) {
                $checkpoint = Get-Content $checkpointPath | ConvertFrom-Json
                $startPosition = $checkpoint.Position
                
                # Verify the checkpoint is for the same file
                if ($checkpoint.FileId -ne $DriveItemId) {
                    # Different file, start fresh
                    $startPosition = 0
                }
            }
            
            # Create or open the destination file
            if ($startPosition -eq 0) {
                # New download, create empty file
                [System.IO.File]::Create($DestinationPath).Close()
            }
            else {
                # Verify the partial file exists and has the expected size
                if (Test-Path $DestinationPath) {
                    $fileInfo = Get-Item $DestinationPath
                    if ($fileInfo.Length -ne $startPosition) {
                        # File size doesn't match checkpoint, start over
                        $startPosition = 0
                        [System.IO.File]::Create($DestinationPath).Close()
                    }
                }
                else {
                    # File missing but we have a checkpoint, start over
                    $startPosition = 0
                    [System.IO.File]::Create($DestinationPath).Close()
                }
            }
            
            # Download the file in chunks
            $currentPosition = $startPosition
            $fileStream = [System.IO.File]::OpenWrite($DestinationPath)
            $fileStream.Seek($currentPosition, [System.IO.SeekOrigin]::Begin) | Out-Null
            
            try {
                while ($currentPosition -lt $FileSize) {
                    $endPosition = [Math]::Min($currentPosition + $chunkSizeBytes - 1, $FileSize - 1)
                    $rangeHeader = "bytes=$currentPosition-$endPosition"
                    $chunkSize = $endPosition - $currentPosition + 1
                    
                    # Get temporary path for this chunk
                    $tempChunkPath = [System.IO.Path]::GetTempFileName()
                    
                    # Download this chunk
                    $attempt = 0
                    $success = $false
                    
                    while (-not $success -and $attempt -lt $MaxRetries) {
                        $attempt++
                        try {
                            # Check rate limit before making request
                            Test-RateLimit
                            
                            $headers = @{
                                "Range" = $rangeHeader
                            }
                            
                            # Use Microsoft Graph API to download the chunk
                            Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/me/drive/items/$DriveItemId/content" -Method GET -Headers $headers -OutputFilePath $tempChunkPath -ErrorAction Stop
                            
                            # Verify the chunk size
                            $chunkInfo = Get-Item $tempChunkPath
                            if ($chunkInfo.Length -eq $chunkSize) {
                                $success = $true
                            }
                            else {
                                throw "Downloaded chunk size ($($chunkInfo.Length)) doesn't match expected size ($chunkSize)"
                            }
                        }
                        catch {
                            if ($attempt -ge $MaxRetries) {
                                throw
                            }
                            
                            # Exponential backoff
                            $delay = [Math]::Pow(2, $attempt)
                            $jitter = Get-Random -Minimum 0 -Maximum 1000
                            $totalDelay = $delay + ($jitter / 1000)
                            Start-Sleep -Seconds $totalDelay
                        }
                    }
                    
                    # Write the chunk to the file
                    $chunkData = [System.IO.File]::ReadAllBytes($tempChunkPath)
                    $fileStream.Write($chunkData, 0, $chunkData.Length)
                    $fileStream.Flush()
                    
                    # Update position and checkpoint
                    $currentPosition = $endPosition + 1
                    $checkpoint = @{
                        FileId = $DriveItemId
                        Position = $currentPosition
                        TotalSize = $FileSize
                        LastUpdated = (Get-Date).ToString('o')
                    }
                    $checkpoint | ConvertTo-Json | Set-Content -Path $checkpointPath
                    
                    # Clean up temp file
                    Remove-Item -Path $tempChunkPath -Force -ErrorAction SilentlyContinue
                    
                    # Report progress
                    $percentComplete = [Math]::Round(($currentPosition / $FileSize) * 100, 1)
                    Write-Verbose "Downloading large file: $percentComplete% complete ($currentPosition of $FileSize bytes)"
                }
                
                # Download complete, remove checkpoint file
                if (Test-Path $checkpointPath) {
                    Remove-Item -Path $checkpointPath -Force -ErrorAction SilentlyContinue
                }
                
                return $true
            }
            finally {
                # Always close the file stream
                if ($fileStream) {
                    $fileStream.Close()
                    $fileStream.Dispose()
                }
            }
        }
        catch {
            Write-Error "Failed to download large file: $_"
            return $false
        }
    }
    
    # Rate limiting function to prevent API throttling
    function Test-RateLimit {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $false)]
            [int]$MaxRequests = $Script:MaxRequestsPerMinute,
            
            [Parameter(Mandatory = $false)]
            [int]$TimeWindowSeconds = 60
        )
        
        $now = Get-Date
        $timeWindow = [TimeSpan]::FromSeconds($TimeWindowSeconds)
        $cutoffTime = $now.Subtract($timeWindow)
        
        # Add the current timestamp
        [void]$Script:RequestTimestamps.Add($now)
        
        # Remove timestamps older than our window
        for ($i = $Script:RequestTimestamps.Count - 1; $i -ge 0; $i--) {
            if ($Script:RequestTimestamps[$i] -lt $cutoffTime) {
                $Script:RequestTimestamps.RemoveAt($i)
            }
        }
        
        # If we have too many requests, wait until we're under the limit
        if ($Script:RequestTimestamps.Count -gt $MaxRequests) {
            $oldestAllowed = $Script:RequestTimestamps[$Script:RequestTimestamps.Count - $MaxRequests]
            $waitTime = ($oldestAllowed.AddSeconds($TimeWindowSeconds) - $now).TotalMilliseconds
            
            if ($waitTime -gt 0) {
                Write-Verbose "Rate limit reached. Waiting $([Math]::Ceiling($waitTime))ms before next request."
                Start-Sleep -Milliseconds ([Math]::Ceiling($waitTime))
            }
        }
    }

#endregion Onedrive Functions


#region Process Functions

    # Main backup function
    function Backup-OneDriveFolder {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$ExportFolder,
        
            [Parameter(Mandatory = $false)]
            [switch]$Overwrite,
        
            [Parameter(Mandatory = $false)]
            [switch]$OverwriteAll,
        
            [Parameter(Mandatory = $false)]
            [switch]$IncludeSharedItems
        )
    
        # catch block around function
        try {

            #post
            # Ensure export folder exists
            New-FolderWithProgress -FolderPath $ExportFolder
        
            # Get OneDrive root
            if (-not $Script:DriveInfo) {
                throw "Could not retrieve OneDrive information"
            }
        
            Write-Host "Starting backup of OneDrive for $($Script:UserName) ($($Script:UserEmail))" -ForegroundColor Gray
            Write-Host "Drive ID: $($Script:DriveInfo.Id)" -ForegroundColor Gray
            Write-Host ""

            # Option to count files for progress reporting
            Write-Host "Would you like to count all files in OneDrive first? (for accurate progress tracking)" -ForegroundColor Yellow
            Write-Host "Note: This may take several minutes for large accounts" -ForegroundColor Yellow

            $countChoice = Read-Host "Count files first? (Y/N)"
            if ($countChoice.ToUpper() -eq "Y") {
                Write-Host "Counting total files in OneDrive (this may take a while)..." -ForegroundColor Yellow
                $count, $deltaData = Get-OneDriveTotalItems
                if ($count -lt 0) {
                    Write-Host "Failed to count files in OneDrive. Proceeding without progress tracking." -ForegroundColor Yellow
                    $Script:TotalFiles = 0  # Set to 0 to disable percentage-based progress
                } else {
                    $Script:TotalFiles = $count
                    Write-Host "Found $count files to process." -ForegroundColor Green
                }
            } else {
                $Script:TotalFiles = 0  # Proceed without counting
                Write-Host "Proceeding without counting files first. Progress will be shown without percentages." -ForegroundColor Yellow
            }

            # Check if we're running in PowerShell 7 or higher for parallel processing
            $isPowerShell7 = $PSVersionTable.PSVersion.Major -ge 7
        
            # Run with parallel batch processing
            if ($isPowerShell7) {
                Write-Host ""
                Write-Host "Using PowerShell 7 with parallel processing" -ForegroundColor Cyan
                
                # Call PS7 implementation
                Start-BackupOnedriveParallel -ExportFolder $ExportFolder -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll -IncludeSharedItems:$IncludeSharedItems
            }
            # Run with sequentiell batch processing
            else {
                Write-Host ""
                Write-Host "Using PowerShell 5.1 sequential processing" -ForegroundColor Cyan

                # Call PS5 implementation
                Start-BackupOnedriveSequential -ExportFolder $ExportFolder -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll -IncludeSharedItems:$IncludeSharedItems
            }
        
            # Complete the progress bar
            Write-Host "" # Clear the spinner line
        
            # Report final statistics
            $elapsedTime = (Get-Date) - $global:StartTime
            $elapsedText = "{0:hh\:mm\:ss}" -f $elapsedTime
            Write-Host "OneDrive backup completed in $elapsedText."
            Write-Host "Summary:"
            Write-Host "  - Total files: $Script:TotalFiles"
            Write-Host "  - Files processed: $global:syncHash.ProcessedFiles"
            Write-Host "  - Files skipped: $global:syncHash.SkippedFiles"
            Write-Host "  - Errors: $global:syncHash.ErrorCount"
            Write-Host "  - Output location: $ExportFolder"
        
            return $true
        }
        catch {
            Write-Host "" # Clear the spinner line
            Write-Error -Exception ([System.Management.Automation.PSInvalidOperationException]::new(
                "Backup failed: $_",
                $_.Exception
            ))
            return $false
        }
    }

    # Function to start a full backup
    function Start-FullBackup {
        Clear-Host
        Write-Host "=======================================================" -ForegroundColor Cyan
        Write-Host "                ONEDRIVE FULL BACKUP                   " -ForegroundColor Cyan
        Write-Host "=======================================================" -ForegroundColor Cyan
        Write-Host ""
    
        # Ensure we're connected
        if (-not $Script:IsConnected) {
            if (-not (Connect-ToOneDrive)) {
                Write-Host "Unable to connect to OneDrive. Returning to main menu..." -ForegroundColor Red
                Start-Sleep -Seconds 3
                Show-WelcomeMenu
                return
            }
            else {
                Write-Host ""
            }
        }
    
        # Ask for export folder
        Write-Host "Please specify where to save your OneDrive backup:" -ForegroundColor Yellow
        $exportFolder = Read-Host "Export Folder Path"
    
        # Validate export folder
        if (-not $exportFolder) {
            Write-Host "No export folder specified. Returning to main menu..." -ForegroundColor Red
            Start-Sleep -Seconds 3
            Show-WelcomeMenu
            return
        }
        
        # Ask about overwriting files
        $validChoice = $false
        $overwrite = $false
        $overwriteAll = $false

        while (-not $validChoice) {
            Write-Host ""
            Write-Host "Do you want to overwrite existing files?" -ForegroundColor Yellow
            Write-Host "Y = Yes when older, N = No for each file, A = Yes to All" -ForegroundColor Gray
            $overwriteChoice = Read-Host "Please choose [ Yes / No / All ]"
            
            switch ($overwriteChoice.ToUpper()) {
                "Y" { 
                    $overwrite = $true
                    $validChoice = $true
                }
                "N" { 
                    $overwrite = $false
                    $validChoice = $true
                }
                "A" { 
                    $overwriteAll = $true
                    $validChoice = $true
                }
                default {
                    Write-Host "Invalid choice. Please enter Y, N, or A." -ForegroundColor Red
                }
            }
        }
    
        # Ask about excluding shared items
        Write-Host ""
        Write-Host "Do you want to include shared items (files/folders shared with you)?" -ForegroundColor Yellow
        $includeSharedChoice = Read-Host "Include shared items? (Y/N)"
        $includeShared = $includeSharedChoice.ToUpper() -eq "Y"
        
        # Reset script variables
        $Script:OverwriteAll = $overwriteAll
        $Script:TotalFiles = 0
        $global:syncHash.SkippedFiles = 0
        $global:syncHash.ErrorCount = 0
        $global:StartTime = Get-Date

        # Start backup
        Write-Host ""
        Write-Host "Starting backup process..." -ForegroundColor Yellow
        Write-Host ""

        # Show exclusions
        if ($Script:ExcludedFolders.Count -gt 0) {
            Write-Host "The following folders will be excluded from backup:" -ForegroundColor Yellow
            foreach ($exclusion in $Script:ExcludedFolders) {
                Write-Host "  - $exclusion" -ForegroundColor Yellow
            }
            Write-Host ""
        }

        if (-not $includeShared) {
            Write-Host "Shared items will be excluded from backup" -ForegroundColor Yellow
            Write-Host ""
        }

        $backupSuccess = Backup-OneDriveFolder -ExportFolder $exportFolder -Overwrite:$overwrite -OverwriteAll:$overwriteAll -IncludeSharedItems:$includeShared
    
        if ($backupSuccess) {
            Write-Host "Backup completed successfully!" -ForegroundColor Green
            Write-Host "Would you like to verify the backup now?" -ForegroundColor Yellow
            $verifyChoice = Read-Host "Verify backup? (Y/N)"
        
            if ($verifyChoice.ToUpper() -eq "Y") {
                # Verify the backup
                Test-OneDriveBackup -BackupFolder $exportFolder -IncludeSharedItems:$includeShared
            }
        }
        else {
            Write-Host ""
            Write-Host "Backup process completed with errors." -ForegroundColor Red
            Write-Host "Please check the error messages above." -ForegroundColor Red
        }
    
        Write-Host ""
        Write-Host "Press any key to return to the main menu..."
        [void][System.Console]::ReadKey($true)
        Show-WelcomeMenu
    }

    # Function to start backup verification
    function Start-BackupVerification {
        Clear-Host
        Write-Host "=======================================================" -ForegroundColor Cyan
        Write-Host "                BACKUP VERIFICATION                    " -ForegroundColor Cyan
        Write-Host "=======================================================" -ForegroundColor Cyan
        Write-Host ""
    
        # Ensure we're connected
        if (-not $Script:IsConnected) {
            if (-not (Connect-ToOneDrive)) {
                Write-Host "Unable to connect to OneDrive. Returning to main menu..." -ForegroundColor Red
                Start-Sleep -Seconds 3
                Show-WelcomeMenu
                return
            }
        }
    
        # Ask for backup folder
        Write-Host "Please specify the location of your OneDrive backup:" -ForegroundColor Yellow
        $backupFolder = Read-Host "Backup Folder Path"
    
        # Validate backup folder
        if (-not $backupFolder -or -not (Test-Path $backupFolder)) {
            Write-Host "Invalid backup folder. Returning to main menu..." -ForegroundColor Red
            Start-Sleep -Seconds 3
            Show-WelcomeMenu
            return
        }
    
        # Verify the backup
        Test-OneDriveBackup -BackupFolder $backupFolder
    
        Write-Host ""
        Write-Host "Press any key to return to the main menu..."
        [void][System.Console]::ReadKey($true)
        Show-WelcomeMenu
    }

    # Function to create a folder with inline progress
    function New-FolderWithProgress {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$FolderPath
        )
        
        if (-not (Test-Path $FolderPath)) {
            New-Item -Path $FolderPath -ItemType Directory -Force | Out-Null
            #DEBUG Write-Host "`rCreated folder: $FolderPath" -NoNewline
        }
    }

    # Function to download a file
    function Save-OneDriveFile {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$DriveItemId,
        
            [Parameter(Mandatory = $true)]
            [string]$DestinationPath,
            
            [Parameter(Mandatory = $false)]
            [datetime]$LastModifiedDateTime,
        
            [Parameter(Mandatory = $false)]
            [switch]$Overwrite,
        
            [Parameter(Mandatory = $false)]
            [switch]$OverwriteAll,
        
            [Parameter(Mandatory = $false)]
            [int]$RetryCount = 3,
        
            [Parameter(Mandatory = $false)]
            [int]$RetryDelaySeconds = 2
        )
    
        try {
            # Check if we should overwrite
            if (Test-Path $DestinationPath) {
                # If it's a folder with the same name as our file, remove it
                if (Test-Path -Path $DestinationPath -PathType Container) {
                    Remove-Item -Path $DestinationPath -Force -Recurse
                    Write-Verbose "Removed folder that had the same name as file: $DestinationPath"
                }
                else {
                    $shouldOverwrite = Test-ShouldOverwriteFile -FilePath $DestinationPath -RemoteLastModified $LastModifiedDateTime -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll
                    if (-not $shouldOverwrite) {
                        Write-Verbose "Skipping file: $DestinationPath"
                        Update-BackupProgress -FileSkipped
                        return
                    }
                }
            }
        
            # Ensure directory exists
            $directory = Split-Path -Path $DestinationPath -Parent
            Write-Host "`rFILE" -ForegroundColor Yellow -NoNewline #DEBUG 
            New-FolderWithProgress -FolderPath $directory
        
            # Download file with retry logic
            $attempt = 0
            $success = $false
            $fileName = Split-Path -Path $DestinationPath -Leaf
        
            while (-not $success -and $attempt -lt $RetryCount) {
                try {
                    $attempt++                    
                    $fileName = Split-Path -Path $DestinationPath -Leaf
                    Write-Host "`rDownloading: $fileName" -NoNewline

                    # Use a temp file for download to avoid partial downloads
                    $tempFile = [System.IO.Path]::GetTempFileName()

                    # Download the file content
                    Get-MgDriveItemContent -DriveId $($Script:DriveInfo.Id) -DriveItemId $DriveItemId -OutFile $tempFile -ErrorAction Stop
                    
                    # Verify the file was created correctly
                    if (Test-Path -Path $tempFile -PathType Container) {
                        throw "File was created as a directory instead of a file"
                    }
                    
                    # Move temp file to destination
                    Move-Item -Path $tempFile -Destination $DestinationPath -Force -ErrorAction Stop

                    # If we made it here, the download was successful
                    $success = $true
                    Update-BackupProgress -FileProcessed
                }
                catch {
                    $lastError = $_

                    # Clean up temp file if it exists
                    if (Test-Path $tempFile) {
                        Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
                    }
                    
                    if ($attempt -ge $RetryCount) {
                        throw
                    }
                
                    # Exponential backoff with jitter
                    $delay = $RetryDelaySeconds * [Math]::Pow(2, $attempt - 1)
                    $jitter = Get-Random -Minimum 0 -Maximum 1000
                    $totalDelay = $delay + ($jitter / 1000)
                    
                    Write-Verbose "Retry $attempt/$RetryCount for $fileName after $totalDelay seconds. Error: $_"
                    Start-Sleep -Milliseconds ($totalDelay * 1000)
                }
            }
        }
        catch {
            Write-Host "" # Clear the line
            Update-BackupProgress -Error

            # Combine with previous error if available
            $errorMessage = if ($lastError) {
                "Failed to download file after $RetryCount attempts. Last error: $lastError"
            } else {
                "Failed to download file: $_"
            }
            
            Write-Error -Exception ([System.IO.IOException]::new(
                "Failed to download file $DriveItemId to $($DestinationPath): $errorMessage",
                $_.Exception
            ))
        }
    }

    # Function to process a folder
    function Save-OneDriveFolder {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$DriveItemId,
        
            [Parameter(Mandatory = $true)]
            [string]$DestinationPath,

            [Parameter(Mandatory = $true)]
            [string]$FolderName,

            [Parameter(Mandatory = $false)]
            [string]$ParentPath = "",
        
            [Parameter(Mandatory = $false)]
            [switch]$Overwrite,
        
            [Parameter(Mandatory = $false)]
            [switch]$OverwriteAll
        )
    
        try {
            $relativePath = if ($ParentPath) { "$ParentPath/$FolderName" } else { $FolderName }
            $relativePath = $relativePath.Replace('\', '/')
            
            # Check if this folder should be excluded
            if (Test-ShouldExcludePath -Path $relativePath) {
                Write-Host "`rSkipping excluded folder: $relativePath" -ForegroundColor Yellow -NoNewline
                return
            }
            
            # Create folder if it doesn't exist
            $folderPath = Join-Path -Path $DestinationPath -ChildPath $FolderName
            Write-Host "`rFOLDER" -ForegroundColor Yellow -NoNewline #DEBUG 
            New-FolderWithProgress -FolderPath $folderPath
        
            # Get items in the folder
            $children = Get-MgDriveItemChild -DriveId $($Script:DriveInfo.Id) -DriveItemId $DriveItemId -ErrorAction Stop
        
            # Process each item
            foreach ($item in $children) {
                $itemRelativePath = "$relativePath/$($item.Name)".Replace('\', '/')
            
                # Skip excluded items
                if (Test-ShouldExcludePath -Path $itemRelativePath) {
                    Write-Host "`rSkipping excluded item: $itemRelativePath" -ForegroundColor Yellow -NoNewline
                    continue
                }
                
                $itemPath = Join-Path -Path $folderPath -ChildPath $item.Name

                if ($item.Folder) {
                    # Process subfolder recursively
                    Save-OneDriveFolder -DriveItemId $item.Id -DestinationPath $folderPath -FolderName $item.Name -ParentPath $relativePath -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll
                }
                else {
                    # Download file
                    Save-OneDriveFile -DriveItemId $item.Id -DestinationPath $itemPath -LastModifiedDateTime $item.LastModifiedDateTime -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll
                }
            }
        }
        catch {
            Write-Host "" # Clear the line
            Update-BackupProgress -Error
            Write-Error -Exception ([System.IO.IOException]::new(
                "Failed to process folder $($FolderName): $_",
                $_.Exception
            ))
        }
    }

    # Function for parallel folder processing (PowerShell 7+ only)
    function Save-OneDriveFolderParallel {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$DriveItemId,
            
            [Parameter(Mandatory = $true)]
            [string]$DestinationPath,
            
            [Parameter(Mandatory = $true)]
            [string]$FolderName,
            
            [Parameter(Mandatory = $false)]
            [switch]$Overwrite,
            
            [Parameter(Mandatory = $false)]
            [switch]$OverwriteAll,
        
            [Parameter(Mandatory = $false)]
            [switch]$IncludeSharedItems
        )
        
        try {
            # Create folder if it doesn't exist
            $folderPath = Join-Path -Path $DestinationPath -ChildPath $FolderName
            New-FolderWithProgress -FolderPath $folderPath
            
            # Get items in the folder
            $items = Get-OneDriveItems -FolderId $DriveItemId -Path $FolderName
            
            # Filter out shared items if not explicitly included
            if (-not $IncludeSharedItems) {
                $originalCount = $items.Count
                $items = $items | Where-Object { -not $_.IsShared }
                $skippedCount = $originalCount - $items.Count
                if ($skippedCount -gt 0) {
                    Write-Verbose "Excluded $skippedCount shared items from folder $FolderName"
                }
            }
            
            # Separate folders and files
            $folders = $items | Where-Object { $_.Type -eq "Folder" }
            $files = $items | Where-Object { $_.Type -eq "File" }
            
            # Process files in parallel
            if ($files.Count -gt 0) {
                # Create script blocks that can be passed to parallel processing
                $SaveFileScriptBlock = ${function:Save-OneDriveFile}
                $TestShouldExcludePathScriptBlock = ${function:Test-ShouldExcludePath}
                
                $files | ForEach-Object -Parallel {
                    $item = $_
                    $itemPath = Join-Path -Path $using:folderPath -ChildPath $item.Name
                    
                    # Skip excluded files
                    if (& $using:TestShouldExcludePathScriptBlock -Path $item.Path) {
                        return
                    }

                    Update-ThrottleLimit
                    
                    # Download the file
                    & $using:SaveFileScriptBlock -DriveItemId $item.Id -DestinationPath $itemPath -LastModifiedDateTime $item.LastModifiedDateTime -Overwrite:$using:Overwrite -OverwriteAll:$using:OverwriteAll
                } -ThrottleLimit $Script:ThrottleLimit
            }
            
            # Process folders sequentially to avoid too many concurrent connections
            foreach ($folder in $folders) {
                # Skip excluded folders
                if (Test-ShouldExcludePath -Path $folder.Path) {
                    Write-Verbose "Skipping excluded folder: $($folder.Path)"
                    continue
                }
                
                # Process folder recursively
                Save-OneDriveFolderParallel -DriveItemId $folder.Id -DestinationPath $folderPath -FolderName $folder.Name -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll -IncludeSharedItems:$IncludeSharedItems
            }
        }
        catch {
            Write-Host "" # Clear the line
            $global:syncHash.ErrorCount++
            Write-Error -Exception ([System.IO.IOException]::new(
                "Failed to process folder $($FolderName): $_",
                $_.Exception
            ))
        }
    }

    # Helper function to extract path from delta item
    function Get-ItemPathFromDelta {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            $Item
        )
        
        # If no parent reference, it's a root item
        if (-not $Item.parentReference -or -not $Item.parentReference.path) {
            return $Item.name
        }
        
        # Extract the path from parent reference (usually in format "/drive/root:/path")
        $parentPath = $Item.parentReference.path
        if ($parentPath -match "root:(.*)") {
            $parentPath = $matches[1].TrimStart('/')
            if ($parentPath) {
                return "$parentPath/$($Item.name)"
            } else {
                return $Item.name
            }
        }
        
        # Fallback - just return the name if we can't parse the path
        return $Item.name
    }

    # PowerShell 7+ implementation with parallel processing
    function Start-BackupOnedriveParallel{
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$ExportFolder,
            
            [Parameter(Mandatory = $false)]
            [switch]$Overwrite,
            
            [Parameter(Mandatory = $false)]
            [switch]$OverwriteAll,
            
            [Parameter(Mandatory = $false)]
            [switch]$IncludeSharedItems
        )
        
        Write-Host "Maximum concurrent operations: $($Script:QueueThrottleLimit)" -ForegroundColor Gray
        
        # Get root items
        Write-Host "Retrieving root items..." -ForegroundColor Yellow
        $rootItems = Get-OneDriveItems
        
        # Filter out shared items if not explicitly included
        if (-not $IncludeSharedItems) {
            Write-Verbose -Message "Excluding all shared objects"
            $originalCount = $rootItems.Count
            $rootItems = $rootItems | Where-Object { -not $_.IsShared }
            $skippedCount = $originalCount - $rootItems.Count
            if ($skippedCount -gt 0) {
                Write-Host "Excluded $skippedCount shared items from root level" -ForegroundColor Yellow
            }
        }
        
        # Separate folders and files
        $rootFolders = $rootItems | Where-Object { $_.Type -eq "Folder" }
        $rootFiles = $rootItems | Where-Object { $_.Type -eq "File" }
        
        Write-Host "Found $($rootFolders.Count) folders and $($rootFiles.Count) files at root level" -ForegroundColor Gray

        # Process root files with batch function
        if ($rootFiles.Count -gt 0) {
            Write-Host "Processing $($rootFiles.Count) root files in batches..." -ForegroundColor Yellow
            
            # Process files in batches using new function
            Save-OneDriveFileBatch -FileItems $rootFiles -DestinationRootPath $ExportFolder -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll
        }
        
        # Process folders recursively
        if ($rootFolders.Count -gt 0) {
            Write-Host "Processing $($rootFolders.Count) root folders..." -ForegroundColor Yellow
            
            # Filter out excluded folders
            $foldersToProcess = $rootFolders | Where-Object { 
                -not (Test-ShouldExcludePath -Path $_.Name)
            }
            
            if ($foldersToProcess.Count -ne $rootFolders.Count) {
                Write-Host "Excluded $($rootFolders.Count - $foldersToProcess.Count) folders due to exclusion rules" -ForegroundColor Yellow
            }
            
            # Process each folder recursively
            foreach ($folder in $foldersToProcess) {
                Write-Host "Processing folder: $($folder.Name)" -ForegroundColor Cyan
                
                # Create the folder structure
                $folderPath = Join-Path -Path $ExportFolder -ChildPath $folder.Name
                New-FolderWithProgress -FolderPath $folderPath
                
                # Process folder contents recursively
                Backup-FolderStructure -FolderId $folder.Id -FolderPath $folder.Path -DestinationPath $folderPath `
                    -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll -IncludeSharedItems:$IncludeSharedItems
            }
        }
    }

    # PowerShell 5.1 implementation with sequential processing
    function Start-BackupOnedriveSequential {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$ExportFolder,
            
            [Parameter(Mandatory = $false)]
            [switch]$Overwrite,
            
            [Parameter(Mandatory = $false)]
            [switch]$OverwriteAll,
            
            [Parameter(Mandatory = $false)]
            [switch]$IncludeSharedItems
        )
        
        # Initialize tracking variables for this specific operation
        $processedFolders = 0
        $failedFolders = 0
        
        # Get all root items
        Write-Host "Retrieving root items..." -ForegroundColor Yellow
        $rootItems = Get-OneDriveItems
        
        # Filter shared items if needed
        if (-not $IncludeSharedItems) {
            $rootItems = $rootItems | Where-Object { -not $_.IsShared }
        }
        
        # Separate folders and files
        $rootFolders = $rootItems | Where-Object { $_.Type -eq "Folder" }
        $rootFiles = $rootItems | Where-Object { $_.Type -eq "File" }
        
        Write-Host "Found $($rootFolders.Count) folders and $($rootFiles.Count) files at root level" -ForegroundColor Gray
        
        # Process root files first
        if ($rootFiles.Count -gt 0) {
            Write-Host "Processing root files..." -ForegroundColor Yellow
            
            # Filter excluded files
            $filesToProcess = $rootFiles | Where-Object {
                -not (Test-ShouldExcludePath -Path $_.Name)
            }
            
            # Process files in batches but without parallelism
            foreach ($file in $filesToProcess) {
                $destinationPath = Join-Path -Path $ExportFolder -ChildPath $file.Name
                
                # Check if we should overwrite
                if (Test-Path $destinationPath) {
                    $shouldOverwrite = Test-ShouldOverwriteFile -FilePath $destinationPath -RemoteLastModified $file.LastModifiedDateTime `
                        -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll
                    
                    if (-not $shouldOverwrite) {
                        Update-BackupProgress -FileSkipped
                        continue
                    }
                }
                
                # Download the file
                Save-OneDriveFile -DriveItemId $file.Id -DestinationPath $destinationPath `
                    -LastModifiedDateTime $file.LastModifiedDateTime -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll
            }
        }
        
        # Build a queue of folders to process
        $foldersToProcess = New-Object System.Collections.Queue
        foreach ($folder in $rootFolders) {
            if (-not (Test-ShouldExcludePath -Path $folder.Name)) {
                $foldersToProcess.Enqueue(@{
                    Id = $folder.Id
                    Name = $folder.Name
                    Path = $folder.Name
                    DestinationPath = $ExportFolder
                })
            } else {
                Write-Host "Skipping excluded folder: $($folder.Name)" -ForegroundColor Yellow
            }
        }
        
        # Process folders in breadth-first order
        $folderCount = $foldersToProcess.Count
        $processedCount = 0
        
        while ($foldersToProcess.Count -gt 0) {
            $folder = $foldersToProcess.Dequeue()
            $processedCount++
            
            # Create folder
            $folderPath = Join-Path -Path $folder.DestinationPath -ChildPath $folder.Name
            New-FolderWithProgress -FolderPath $folderPath
            
            # Show progress
            $percentComplete = if ($folderCount -gt 0) { [math]::Round(($processedCount / $folderCount) * 100) } else { 0 }
            Show-SpinnerProgress -Activity "Processing folders" -Status "Folder $processedCount ($percentComplete%)" -PercentComplete $percentComplete -NoNewLine
            
            try {
                # Get items in this folder
                $folderItems = Get-OneDriveItems -FolderId $folder.Id -Path $folder.Path
                
                # Filter shared items if needed
                if (-not $IncludeSharedItems) {
                    $folderItems = $folderItems | Where-Object { -not $_.IsShared }
                }
                
                # Process files in this folder
                $files = $folderItems | Where-Object { $_.Type -eq "File" }
                $filesToProcess = $files | Where-Object { -not (Test-ShouldExcludePath -Path $_.Path) }
                
                # Process files
                foreach ($file in $filesToProcess) {
                    $filePath = Join-Path -Path $folderPath -ChildPath $file.Name
                    
                    # Check if we should overwrite
                    if (Test-Path $filePath) {
                        $shouldOverwrite = Test-ShouldOverwriteFile -FilePath $filePath -RemoteLastModified $file.LastModifiedDateTime `
                            -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll
                        
                        if (-not $shouldOverwrite) {
                            Update-BackupProgress -FileSkipped
                            continue
                        }
                    }
                    
                    # Download the file
                    Save-OneDriveFile -DriveItemId $file.Id -DestinationPath $filePath `
                        -LastModifiedDateTime $file.LastModifiedDateTime -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll
                }
                
                # Queue subfolders for processing
                $subfolders = $folderItems | Where-Object { $_.Type -eq "Folder" }
                foreach ($subfolder in $subfolders) {
                    if (-not (Test-ShouldExcludePath -Path $subfolder.Path)) {
                        $foldersToProcess.Enqueue(@{
                            Id = $subfolder.Id
                            Name = $subfolder.Name
                            Path = $subfolder.Path
                            DestinationPath = $folderPath
                        })
                        $folderCount++
                    }
                }
                
                $processedFolders++
            }
            catch {
                Write-Warning "Failed to process folder $($folder.Path): $_"
                $failedFolders++
                Update-BackupProgress -Error
            }
        }
    # Complete the progress
    Write-Host "" # Clear the spinner line
        
    # Display summary
    Write-Host ""
    Write-Host "Sequential processing completed!" -ForegroundColor Green
    Write-Host "Summary:" -ForegroundColor Yellow
    Write-Host "  Folders processed: $processedFolders (Failed: $failedFolders)"
    Write-Host "  Files processed: $($global:syncHash.ProcessedFiles.Count)"
    Write-Host "  Files skipped: $($global:syncHash.SkippedFiles)"
    Write-Host "  Errors: $($global:syncHash.ErrorCount)"
    }

#endregion Process Functions


#region Helper-Functions

    # Function to update progress
    function Update-BackupProgress {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $false)]
            [switch]$FileProcessed,
        
            [Parameter(Mandatory = $false)]
            [switch]$FileSkipped,
        
            [Parameter(Mandatory = $false)]
            [switch]$Error
        )

        if ($FileProcessed) {
            # Use lock for atomic increment
            [System.Threading.Monitor]::Enter($global:ProcessedFilesLock)
            try {
                $global:syncHash.ProcessedFiles++
                # Also update Statistics if it exists
                if ($null -ne $global:Statistics) {
                    $global:Statistics.AddOrUpdate("ProcessedFiles", 1, [func[string,int,int]]{
                        param($key, $oldVal) $oldVal + 1
                    })
                }
            }
            finally {
                [System.Threading.Monitor]::Exit($global:ProcessedFilesLock)
            }
        }

        if ($FileSkipped) {
            # Use lock for atomic increment
            [System.Threading.Monitor]::Enter($global:SkippedFilesLock)
            try {
                $global:syncHash.SkippedFiles++
                # Also update Statistics if it exists
                if ($null -ne $global:Statistics) {
                    $global:Statistics.AddOrUpdate("SkippedFiles", 1, [func[string,int,int]]{
                        param($key, $oldVal) $oldVal + 1
                    })
                }
            }
            finally {
                [System.Threading.Monitor]::Exit($global:SkippedFilesLock)
            }
        }

        if ($Error) {
            # Use lock for atomic increment
            [System.Threading.Monitor]::Enter($global:ErrorCountLock)
            try {
                $global:syncHash.ErrorCount++
                # Also update Statistics if it exists
                if ($null -ne $global:Statistics) {
                    $global:Statistics.AddOrUpdate("ErrorCount", 1, [func[string,int,int]]{
                        param($key, $oldVal) $oldVal + 1
                    })
                }
            }
            finally {
                [System.Threading.Monitor]::Exit($global:ErrorCountLock)
            }
        }

        # Get values in a thread-safe way
        [System.Threading.Monitor]::Enter($global:ProcessedFilesLock)
        try { $processedFiles = $global:syncHash.ProcessedFiles } 
        finally { [System.Threading.Monitor]::Exit($global:ProcessedFilesLock) }
        
        [System.Threading.Monitor]::Enter($global:SkippedFilesLock)
        try { $skippedFiles = $global:syncHash.SkippedFiles } 
        finally { [System.Threading.Monitor]::Exit($global:SkippedFilesLock) }
        
        [System.Threading.Monitor]::Enter($global:ErrorCountLock)
        try { $errorCount = $global:syncHash.ErrorCount } 
        finally { [System.Threading.Monitor]::Exit($global:ErrorCountLock) }
        
        $totalProcessed = $processedFiles + $skippedFiles
        if ($totalProcessed % 10 -eq 0) {
            $elapsedTime = (Get-Date) - $global:StartTime
            $elapsedText = "{0:hh\:mm\:ss}" -f $elapsedTime
            
            if ($Script:TotalFiles -gt 0) {
                $percentage = [math]::Min(100, [math]::Round(($totalProcessed / $Script:TotalFiles) * 100, 1))
                Show-SpinnerProgress -Activity "Backing up OneDrive" -Status "$totalProcessed/$Script:TotalFiles files" -PercentComplete $percentage -NoNewLine
                Write-Host "`rProgress: $percentage% ($totalProcessed/$Script:TotalFiles) | Elapsed: $elapsedText | Processed: $processedFiles | Skipped: $skippedFiles | Errors: $errorCount" -NoNewline
            } else {
                # If we don't know the total, just show counts without percentages
                Show-SpinnerProgress -Activity "Backing up OneDrive" -Status "$totalProcessed/$Script:TotalFiles files" -PercentComplete -1 -NoNewLine
                Write-Host "`rProgress: Processed $totalProcessed files | Elapsed: $elapsedText | Processed: $processedFiles | Skipped: $skippedFiles | Errors: $errorCount" -NoNewline
            }
        }
    }

    # Function to check if a file should be overwritten
    function Test-ShouldOverwriteFile {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$FilePath,
        
            [Parameter(Mandatory = $false)]
            [datetime]$RemoteLastModified,
        
            [Parameter(Mandatory = $false)]
            [switch]$Overwrite,
        
            [Parameter(Mandatory = $false)]
            [switch]$OverwriteAll
        )
    
        if (-not (Test-Path $FilePath)) {
            return $true
        }
    
        if ($Script:OverwriteAll -or $OverwriteAll) {
            return $true
        }
    
        if ($Overwrite) {
            $existingFile = Get-Item $FilePath
            $existingLastWrite = $existingFile.LastWriteTime.ToUniversalTime()
            
            # Compare timestamps if remote timestamp is provided
            if ($RemoteLastModified) {
                # Return true if remote file is newer (overwrite local file)
                if ($RemoteLastModified -gt $existingLastWrite) {
                    Write-Verbose "Remote file ($RemoteLastModified) is newer than local file ($existingLastWrite). Overwriting."
                    return $true
                } else {
                    Write-Verbose "Local file ($existingLastWrite) is same or newer than remote file ($RemoteLastModified). Skipping."
                    return $false
                }
            }
            
            # Default to overwrite if no timestamp comparison is possible
            Write-Warning -Message "No timestamp comparison possible"
            return $false
        }
    
        # Ask the user
        $response = Read-Host "File '$FilePath' already exists. Overwrite? (Y/N/A for all)"
        if ($response -eq "A") {
            # Set script-level variable to overwrite all
            $Script:OverwriteAll = $true
            return $true
        }
        return $response -eq "Y"
    }

    # Main backup function
    function Backup-OneDriveFolder {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$ExportFolder,
            
            [Parameter(Mandatory = $false)]
            [switch]$Overwrite,
            
            [Parameter(Mandatory = $false)]
            [switch]$OverwriteAll,
            
            [Parameter(Mandatory = $false)]
            [switch]$IncludeSharedItems
        )
        
        # catch block around function
        try {
            # Ensure export folder exists
            New-FolderWithProgress -FolderPath $ExportFolder
            
            # Get OneDrive root
            if (-not $Script:DriveInfo) {
                throw "Could not retrieve OneDrive information"
            }
            
            Write-Host "Starting backup of OneDrive for $($Script:UserName) ($($Script:UserEmail))" -ForegroundColor Gray
            Write-Host "Drive ID: $($Script:DriveInfo.Id)" -ForegroundColor Gray
            Write-Host ""

            # Option to count files for progress reporting
            Write-Host "Would you like to count all files in OneDrive first? (for accurate progress tracking)" -ForegroundColor Yellow
            Write-Host "Note: This may take several minutes for large accounts" -ForegroundColor Yellow

            $countChoice = Read-Host "Count files first? (Y/N)"
            if ($countChoice.ToUpper() -eq "Y") {
                Write-Host "Counting total files in OneDrive (this may take a while)..." -ForegroundColor Yellow
                $Script:TotalFiles = Get-OneDriveTotalItems
                if ($Script:TotalFiles -lt 0) {
                    Write-Host "Failed to count files in OneDrive. Proceeding without progress tracking." -ForegroundColor Yellow
                    $Script:TotalFiles = 0  # Set to 0 to disable percentage-based progress
                } else {
                    Write-Host "Found $($Script:TotalFiles) files to process." -ForegroundColor Green
                }
            } else {
                $Script:TotalFiles = 0  # Proceed without counting
                Write-Host "Proceeding without counting files first. Progress will be shown without percentages." -ForegroundColor Yellow
            }

            # Check if we're running in PowerShell 7 or higher for parallel processing
            $isPowerShell7 = $PSVersionTable.PSVersion.Major -ge 7
            
            # Run with parallel batch processing
            if ($isPowerShell7) {
                Write-Host ""
                Write-Host "Using PowerShell 7 with parallel processing" -ForegroundColor Cyan
                
                # Call PS7 implementation
                Start-BackupOnedriveParallel -ExportFolder $ExportFolder -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll -IncludeSharedItems:$IncludeSharedItems
            }
            # Run with sequential batch processing
            else {
                Write-Host ""
                Write-Host "Using PowerShell 5.1 sequential processing" -ForegroundColor Cyan

                # Call PS5 implementation
                Start-BackupOnedriveSequential -ExportFolder $ExportFolder -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll -IncludeSharedItems:$IncludeSharedItems
            }
            
            # Complete the progress bar
            Write-Host "" # Clear the spinner line
            
            # Report final statistics
            $elapsedTime = (Get-Date) - $global:StartTime
            $elapsedText = "{0:hh\:mm\:ss}" -f $elapsedTime
            Write-Host "OneDrive backup completed in $elapsedText."
            Write-Host "Summary:"
            Write-Host "  - Total files: $Script:TotalFiles"
            Write-Host "  - Files processed: $global:syncHash.ProcessedFiles"
            Write-Host "  - Files skipped: $global:syncHash.SkippedFiles"
            Write-Host "  - Errors: $global:syncHash.ErrorCount"
            Write-Host "  - Output location: $ExportFolder"
            
            return $true
        }
        catch {
            Write-Host "" # Clear the spinner line
            Write-Error -Exception ([System.Management.Automation.PSInvalidOperationException]::new(
                "Backup failed: $_",
                $_.Exception
            ))
            return $false
        }
    }

    # PowerShell 7+ implementation with parallel processing
    function Start-BackupOnedriveParallel {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$ExportFolder,
            
            [Parameter(Mandatory = $false)]
            [switch]$Overwrite,
            
            [Parameter(Mandatory = $false)]
            [switch]$OverwriteAll,
            
            [Parameter(Mandatory = $false)]
            [switch]$IncludeSharedItems
        )
        
        Write-Host "Maximum concurrent operations: $($Script:QueueThrottleLimit)" -ForegroundColor Gray
        
        # Get root items
        Write-Host "Retrieving root items..." -ForegroundColor Yellow
        $rootItems = Get-OneDriveItems
        
        # Filter out shared items if not explicitly included
        if (-not $IncludeSharedItems) {
            Write-Verbose -Message "Excluding all shared objects"
            $originalCount = $rootItems.Count
            $rootItems = $rootItems | Where-Object { -not $_.IsShared }
            $skippedCount = $originalCount - $rootItems.Count
            if ($skippedCount -gt 0) {
                Write-Host "Excluded $skippedCount shared items from root level" -ForegroundColor Yellow
            }
        }
        
        # Separate folders and files
        $rootFolders = $rootItems | Where-Object { $_.Type -eq "Folder" }
        $rootFiles = $rootItems | Where-Object { $_.Type -eq "File" }
        
        Write-Host "Found $($rootFolders.Count) folders and $($rootFiles.Count) files at root level" -ForegroundColor Gray

        # Process root files with batch function
        if ($rootFiles.Count -gt 0) {
            Write-Host "Processing $($rootFiles.Count) root files in batches..." -ForegroundColor Yellow
            
            # Process files in batches using new function
            Save-OneDriveFileBatch -FileItems $rootFiles -DestinationRootPath $ExportFolder -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll
        }
        
        # Process folders recursively
        if ($rootFolders.Count -gt 0) {
            Write-Host "Processing $($rootFolders.Count) root folders..." -ForegroundColor Yellow
            
            # Filter out excluded folders
            $foldersToProcess = $rootFolders | Where-Object { 
                -not (Test-ShouldExcludePath -Path $_.Name)
            }
            
            if ($foldersToProcess.Count -ne $rootFolders.Count) {
                Write-Host "Excluded $($rootFolders.Count - $foldersToProcess.Count) folders due to exclusion rules" -ForegroundColor Yellow
            }
            
            # Process each folder recursively
            foreach ($folder in $foldersToProcess) {
                Write-Host "Processing folder: $($folder.Name)" -ForegroundColor Cyan
                
                # Create the folder structure
                $folderPath = Join-Path -Path $ExportFolder -ChildPath $folder.Name
                New-FolderWithProgress -FolderPath $folderPath
                
                # Process folder contents recursively
                Backup-FolderStructure -FolderId $folder.Id -FolderPath $folder.Path -DestinationPath $folderPath `
                    -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll -IncludeSharedItems:$IncludeSharedItems
            }
        }
    }

    # PowerShell 5.1 implementation with sequential processing
    function Start-BackupOnedriveSequential {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$ExportFolder,
            
            [Parameter(Mandatory = $false)]
            [switch]$Overwrite,
            
            [Parameter(Mandatory = $false)]
            [switch]$OverwriteAll,
            
            [Parameter(Mandatory = $false)]
            [switch]$IncludeSharedItems
        )
        
        # Initialize tracking variables for this specific operation
        $processedFolders = 0
        $failedFolders = 0
        
        # Get all root items
        Write-Host "Retrieving root items..." -ForegroundColor Yellow
        $rootItems = Get-OneDriveItems
        
        # Filter shared items if needed
        if (-not $IncludeSharedItems) {
            $rootItems = $rootItems | Where-Object { -not $_.IsShared }
        }
        
        # Separate folders and files
        $rootFolders = $rootItems | Where-Object { $_.Type -eq "Folder" }
        $rootFiles = $rootItems | Where-Object { $_.Type -eq "File" }
        
        Write-Host "Found $($rootFolders.Count) folders and $($rootFiles.Count) files at root level" -ForegroundColor Gray
        
        # Process root files first
        if ($rootFiles.Count -gt 0) {
            Write-Host "Processing root files..." -ForegroundColor Yellow
            
            # Filter excluded files
            $filesToProcess = $rootFiles | Where-Object {
                -not (Test-ShouldExcludePath -Path $_.Name)
            }
            
            # Process files in batches but without parallelism
            foreach ($file in $filesToProcess) {
                $destinationPath = Join-Path -Path $ExportFolder -ChildPath $file.Name
                
                # Check if we should overwrite
                if (Test-Path $destinationPath) {
                    $shouldOverwrite = Test-ShouldOverwriteFile -FilePath $destinationPath -RemoteLastModified $file.LastModifiedDateTime `
                        -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll
                    
                    if (-not $shouldOverwrite) {
                        Update-BackupProgress -FileSkipped
                        continue
                    }
                }
                
                # Download the file
                Save-OneDriveFile -DriveItemId $file.Id -DestinationPath $destinationPath `
                    -LastModifiedDateTime $file.LastModifiedDateTime -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll
            }
        }
        
        # Build a queue of folders to process
        $foldersToProcess = New-Object System.Collections.Queue
        foreach ($folder in $rootFolders) {
            if (-not (Test-ShouldExcludePath -Path $folder.Name)) {
                $foldersToProcess.Enqueue(@{
                    Id = $folder.Id
                    Name = $folder.Name
                    Path = $folder.Name
                    DestinationPath = $ExportFolder
                })
            } else {
                Write-Host "Skipping excluded folder: $($folder.Name)" -ForegroundColor Yellow
            }
        }
        
        # Process folders in breadth-first order
        $folderCount = $foldersToProcess.Count
        $processedCount = 0
        
        while ($foldersToProcess.Count -gt 0) {
            $folder = $foldersToProcess.Dequeue()
            $processedCount++
            
            # Create folder
            $folderPath = Join-Path -Path $folder.DestinationPath -ChildPath $folder.Name
            New-FolderWithProgress -FolderPath $folderPath
            
            # Show progress
            $percentComplete = if ($folderCount -gt 0) { [math]::Round(($processedCount / $folderCount) * 100) } else { 0 }
            Show-SpinnerProgress -Activity "Processing folders" -Status "Folder $processedCount ($percentComplete%)" -PercentComplete $percentComplete -NoNewLine
            
            try {
                # Get items in this folder
                $folderItems = Get-OneDriveItems -FolderId $folder.Id -Path $folder.Path
                
                # Filter shared items if needed
                if (-not $IncludeSharedItems) {
                    $folderItems = $folderItems | Where-Object { -not $_.IsShared }
                }
                
                # Process files in this folder
                $files = $folderItems | Where-Object { $_.Type -eq "File" }
                $filesToProcess = $files | Where-Object { -not (Test-ShouldExcludePath -Path $_.Path) }
                
                # Process files
                foreach ($file in $filesToProcess) {
                    $filePath = Join-Path -Path $folderPath -ChildPath $file.Name
                    
                    # Check if we should overwrite
                    if (Test-Path $filePath) {
                        $shouldOverwrite = Test-ShouldOverwriteFile -FilePath $filePath -RemoteLastModified $file.LastModifiedDateTime `
                            -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll
                        
                        if (-not $shouldOverwrite) {
                            Update-BackupProgress -FileSkipped
                            continue
                        }
                    }
                    
                    # Download the file
                    Save-OneDriveFile -DriveItemId $file.Id -DestinationPath $filePath `
                        -LastModifiedDateTime $file.LastModifiedDateTime -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll
                }
                
                # Queue subfolders for processing
                $subfolders = $folderItems | Where-Object { $_.Type -eq "Folder" }
                foreach ($subfolder in $subfolders) {
                    if (-not (Test-ShouldExcludePath -Path $subfolder.Path)) {
                        $foldersToProcess.Enqueue(@{
                            Id = $subfolder.Id
                            Name = $subfolder.Name
                            Path = $subfolder.Path
                            DestinationPath = $folderPath
                        })
                        $folderCount++
                    }
                }
                
                $processedFolders++
            }
            catch {
                Write-Warning "Failed to process folder $($folder.Path): $_"
                $failedFolders++
                Update-BackupProgress -Error
            }
        }
        # Complete the progress
        Write-Host "" # Clear the spinner line
        
        # Display summary
        Write-Host ""
        Write-Host "Sequential processing completed!" -ForegroundColor Green
        Write-Host "Summary:" -ForegroundColor Yellow
        Write-Host "  Folders processed: $processedFolders (Failed: $failedFolders)"
        Write-Host "  Files processed: $($global:syncHash.ProcessedFiles.Count)"
        Write-Host "  Files skipped: $($global:syncHash.SkippedFiles)"
        Write-Host "  Errors: $($global:syncHash.ErrorCount)"
    }

    function Backup-FolderStructure {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$FolderId,
            
            [Parameter(Mandatory = $true)]
            [string]$FolderPath,
            
            [Parameter(Mandatory = $true)]
            [string]$DestinationPath,
            
            [Parameter(Mandatory = $false)]
            [switch]$Overwrite,
            
            [Parameter(Mandatory = $false)]
            [switch]$OverwriteAll,
            
            [Parameter(Mandatory = $false)]
            [switch]$IncludeSharedItems
        )
        
        # Get all items in the folder
        $folderItems = Get-OneDriveItems -FolderId $FolderId -Path $FolderPath
        
        # Filter shared items if needed
        if (-not $IncludeSharedItems) {
            $folderItems = $folderItems | Where-Object { -not $_.IsShared }
        }
        
        # Separate files and folders
        $files = $folderItems | Where-Object { $_.Type -eq "File" }
        $subFolders = $folderItems | Where-Object { $_.Type -eq "Folder" }
        
        # Process files in batches
        if ($files.Count -gt 0) {
            # Filter excluded files
            $filesToProcess = $files | Where-Object {
                -not (Test-ShouldExcludePath -Path $_.Path)
            }
            
            if ($filesToProcess.Count -gt 0) {
                Save-OneDriveFileBatch -FileItems $filesToProcess -DestinationRootPath $DestinationPath `
                    -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll
            }
        }
        
        # Process subfolders recursively
        foreach ($subFolder in $subFolders) {
            # Skip excluded folders
            if (Test-ShouldExcludePath -Path $subFolder.Path) {
                Write-Verbose "Skipping excluded folder: $($subFolder.Path)"
                continue
            }
            
            # Create subfolder
            $subFolderPath = Join-Path -Path $DestinationPath -ChildPath $subFolder.Name
            New-FolderWithProgress -FolderPath $subFolderPath
            
            # Process subfolder recursively
            Backup-FolderStructure -FolderId $subFolder.Id -FolderPath $subFolder.Path -DestinationPath $subFolderPath `
                -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll -IncludeSharedItems:$IncludeSharedItems
        }
    }
        
#endregion Helper-Functions


#region Main
    try {
        # Initialize modules
        if (-not (Initialize-GraphModules -UpdateModules:$UpdateModules)) {
            throw "Failed to initialize required modules"
        }
    
        # Start with the welcome menu
        Show-WelcomeMenu
    }
    catch {
        Write-Error "Script execution failed: $_"
        exit 1
    }

#endregion Main

}