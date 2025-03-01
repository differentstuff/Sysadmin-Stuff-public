function Backup-OneDrivePersonal{
<#
.SYNOPSIS
    Backs up OneDrive Personal to a local folder.
.DESCRIPTION
    This script backs up all files and folders from OneDrive Personal to a specified local folder.
    It supports both PowerShell 5.1 (sequential processing) and PowerShell 7+ (parallel processing).
.PARAMETER ExportFolder
    The folder where OneDrive files will be exported to.
.PARAMETER UpdateModules
    If specified, updates the Microsoft Graph PowerShell modules.
.PARAMETER Overwrite
    If specified, overwrites existing files that are older than the OneDrive version.
.PARAMETER OverwriteAll
    If specified, overwrites all existing files without prompting.
.EXAMPLE
    .\Backup-OneDrivePersonal.ps1 -ExportFolder "E:\OneDriveBackup"
.EXAMPLE
    .\Backup-OneDrivePersonal.ps1 -ExportFolder "E:\OneDriveBackup" -UpdateModules -OverwriteAll
#>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ExportFolder,
    
        [Parameter(Mandatory = $false)]
        [switch]$UpdateModules,
    
        [Parameter(Mandatory = $false)]
        [switch]$Overwrite,
    
        [Parameter(Mandatory = $false)]
        [switch]$OverwriteAll
    )

    # Script-level variables
    $Script:OverwriteAll = $OverwriteAll
    $Script:TotalFiles = 0
    $Script:ProcessedFiles = 0
    $Script:SkippedFiles = 0
    $Script:ErrorCount = 0
    $Script:StartTime = Get-Date

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

    # Function to connect to Microsoft Graph
    function Connect-OneDriveGraph {
        [CmdletBinding()]
        param()
    
        try {
            # Define required scopes for OneDrive access
            $scopes = @(
                "Files.Read",
                "Files.Read.All",
                "Sites.Read.All",
                "User.Read"
            )
        
            # Try to connect with interactive authentication for personal accounts
            Write-Host "Connecting to Microsoft Graph..."
            Connect-MgGraph -Scopes $scopes -TenantId "consumers" -ErrorAction Stop
        
            # Verify connection
            $context = Get-MgContext
            if (-not $context) {
                throw "Failed to establish Microsoft Graph connection"
            }
        
            Write-Host "Connected to Microsoft Graph as $($context.Account)"
            return $true
        }
        catch {
            Write-Error -Exception ([System.Management.Automation.PSInvalidOperationException]::new(
                "Authentication failed: $_",
                $_.Exception
            ))
            return $false
        }
    }

    # Function to check if a file should be overwritten
    function Test-ShouldOverwriteFile {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$FilePath,
        
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
            $existingLastWrite = $existingFile.LastWriteTime
        
            # If we had the remote file's last write time, we could compare here
            # For now, we'll just use the Overwrite switch
            return $true
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
            $Script:ProcessedFiles++
        }
    
        if ($FileSkipped) {
            $Script:SkippedFiles++
        }
    
        if ($Error) {
            $Script:ErrorCount++
        }
    
        $totalProcessed = $Script:ProcessedFiles + $Script:SkippedFiles
        if ($totalProcessed % 10 -eq 0 -and $Script:TotalFiles -gt 0) {
            $percentage = [math]::Min(100, [math]::Round(($totalProcessed / $Script:TotalFiles) * 100, 1))
            $elapsedTime = (Get-Date) - $Script:StartTime
            $elapsedText = "{0:hh\:mm\:ss}" -f $elapsedTime
        
            Write-Progress -Activity "Backing up OneDrive" -Status "$percentage% Complete ($totalProcessed/$Script:TotalFiles)" -PercentComplete $percentage
            Write-Host "Progress: $percentage% ($totalProcessed/$Script:TotalFiles) | Elapsed: $elapsedText | Processed: $($Script:ProcessedFiles) | Skipped: $($Script:SkippedFiles) | Errors: $($Script:ErrorCount)"
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
            [switch]$Overwrite,
        
            [Parameter(Mandatory = $false)]
            [switch]$OverwriteAll,
        
            [Parameter(Mandatory = $false)]
            [int]$RetryCount = 3
        )
    
        try {
            # Check if we should overwrite
            if (Test-Path $DestinationPath) {
                $shouldOverwrite = Test-ShouldOverwriteFile -FilePath $DestinationPath -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll
                if (-not $shouldOverwrite) {
                    Write-Verbose "Skipping file: $DestinationPath"
                    Update-BackupProgress -FileSkipped
                    return
                }
            }
        
            # Ensure directory exists
            $directory = Split-Path -Path $DestinationPath -Parent
            if (-not (Test-Path $directory)) {
                New-Item -Path $directory -ItemType Directory -Force | Out-Null
            }
        
            # Download file with retry logic
            $attempt = 0
            $success = $false
        
            while (-not $success -and $attempt -lt $RetryCount) {
                try {
                    $attempt++
                    Write-Host "Downloading file to: $DestinationPath"
                    Get-MgDriveItemContent -DriveId (Get-MgDrive).Id -DriveItemId $DriveItemId -OutFile $DestinationPath -ErrorAction Stop
                    $success = $true
                    Update-BackupProgress -FileProcessed
                }
                catch {
                    if ($attempt -ge $RetryCount) {
                        throw
                    }
                
                    Write-Warning "Download attempt $attempt failed for $DestinationPath. Retrying in 5 seconds..."
                    Start-Sleep -Seconds 5
                }
            }
        }
        catch {
            Update-BackupProgress -Error
            Write-Error -Exception ([System.IO.IOException]::new(
                "Failed to download file $DriveItemId to $DestinationPath: $_",
                $_.Exception
            ))
        }
    }

    # Function to count total files in a folder recursively
    function Get-OneDriveTotalItems {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$DriveItemId
        )
    
        try {
            $count = 0
            $children = Get-MgDriveItemChild -DriveId (Get-MgDrive).Id -DriveItemId $DriveItemId -ErrorAction Stop
        
            # Count files
            $count += ($children | Where-Object { -not $_.Folder }).Count
        
            # Recursively count files in subfolders
            foreach ($folder in ($children | Where-Object { $_.Folder })) {
                $count += Get-OneDriveTotalItems -DriveItemId $folder.Id
            }
        
            return $count
        }
        catch {
            Write-Warning "Failed to count items in folder $DriveItemId: $_"
            return 0
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
            [switch]$Overwrite,
        
            [Parameter(Mandatory = $false)]
            [switch]$OverwriteAll
        )
    
        try {
            # Create folder if it doesn't exist
            $folderPath = Join-Path -Path $DestinationPath -ChildPath $FolderName
            if (-not (Test-Path $folderPath)) {
                New-Item -Path $folderPath -ItemType Directory -Force | Out-Null
                Write-Host "Created folder: $folderPath"
            }
        
            # Get items in the folder
            $children = Get-MgDriveItemChild -DriveId (Get-MgDrive).Id -DriveItemId $DriveItemId -ErrorAction Stop
        
            # Process each item
            foreach ($item in $children) {
                $itemPath = Join-Path -Path $folderPath -ChildPath $item.Name
            
                if ($item.Folder) {
                    # Process subfolder recursively
                    Save-OneDriveFolder -DriveItemId $item.Id -DestinationPath $folderPath -FolderName $item.Name -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll
                }
                else {
                    # Download file
                    Save-OneDriveFile -DriveItemId $item.Id -DestinationPath $itemPath -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll
                }
            }
        }
        catch {
            Update-BackupProgress -Error
            Write-Error -Exception ([System.IO.IOException]::new(
                "Failed to process folder $FolderName: $_",
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
            [switch]$OverwriteAll
        )
    
        try {
            # Create folder if it doesn't exist
            $folderPath = Join-Path -Path $DestinationPath -ChildPath $FolderName
            if (-not (Test-Path $folderPath)) {
                New-Item -Path $folderPath -ItemType Directory -Force | Out-Null
                Write-Host "Created folder: $folderPath"
            }
        
            # Get items in the folder
            $children = Get-MgDriveItemChild -DriveId (Get-MgDrive).Id -DriveItemId $DriveItemId -ErrorAction Stop
        
            # Process files in parallel
            $files = $children | Where-Object { -not $_.Folder }
            if ($files.Count -gt 0) {
                $files | ForEach-Object -Parallel {
                    $item = $_
                    $itemPath = Join-Path -Path $using:folderPath -ChildPath $item.Name
                
                    # Using script block to call the function in the parent scope
                    & $using:SaveFileScriptBlock -DriveItemId $item.Id -DestinationPath $itemPath -Overwrite:$using:Overwrite -OverwriteAll:$using:OverwriteAll
                
                    # Update progress (via shared variable)
                    $using:UpdateProgressScriptBlock.Invoke()
                } -ThrottleLimit 5
            }
        
            # Process folders sequentially to avoid too many concurrent connections
            $folders = $children | Where-Object { $_.Folder }
            foreach ($folder in $folders) {
                Save-OneDriveFolderParallel -DriveItemId $folder.Id -DestinationPath $folderPath -FolderName $folder.Name -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll
            }
        }
        catch {
            $Script:ErrorCount++
            Write-Error -Exception ([System.IO.IOException]::new(
                "Failed to process folder $FolderName: $_",
                $_.Exception
            ))
        }
    }

    # Main backup function
    function Backup-OneDrivePersonal {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$ExportFolder,
        
            [Parameter(Mandatory = $false)]
            [switch]$Overwrite,
        
            [Parameter(Mandatory = $false)]
            [switch]$OverwriteAll
        )
    
        try {
            # Ensure export folder exists
            if (-not (Test-Path $ExportFolder)) {
                New-Item -Path $ExportFolder -ItemType Directory -Force | Out-Null
                Write-Host "Created export folder: $ExportFolder"
            }
        
            # Get OneDrive root
            $drive = Get-MgDrive -ErrorAction Stop
            if (-not $drive) {
                throw "Could not retrieve OneDrive information"
            }
        
            Write-Host "Starting backup of OneDrive: $($drive.Name)"
        
            # Get root items
            $rootItems = Get-MgDriveRoot -DriveId $drive.Id | Get-MgDriveRootChildren -ErrorAction Stop
        
            # Count total files for progress reporting
            Write-Host "Counting total files in OneDrive (this may take a while)..."
            $Script:TotalFiles = 0
            foreach ($item in $rootItems) {
                if ($item.Folder) {
                    $Script:TotalFiles += Get-OneDriveTotalItems -DriveItemId $item.Id
                }
                else {
                    $Script:TotalFiles++
                }
            }
            Write-Host "Found $($Script:TotalFiles) files to process"
        
            # Check if we're running in PowerShell 7 or higher for parallel processing
            $isPowerShell7 = $PSVersionTable.PSVersion.Major -ge 7
        
            if ($isPowerShell7) {
                Write-Host "Using PowerShell 7+ parallel processing"
            
                # Create script blocks for parallel processing
                $SaveFileScriptBlock = ${function:Save-OneDriveFile}
                $UpdateProgressScriptBlock = {
                    $Script:ProcessedFiles++
                    $totalProcessed = $Script:ProcessedFiles + $Script:SkippedFiles
                    if ($totalProcessed % 10 -eq 0) {
                        $percentage = [math]::Min(100, [math]::Round(($totalProcessed / $Script:TotalFiles) * 100, 1))
                        Write-Host "Progress: $percentage% ($totalProcessed/$Script:TotalFiles)"
                    }
                }
            
                # Process folders (we'll handle these one by one to avoid too many concurrent connections)
                $folders = $rootItems | Where-Object { $_.Folder }
                foreach ($folder in $folders) {
                    Save-OneDriveFolderParallel -DriveItemId $folder.Id -DestinationPath $ExportFolder -FolderName $folder.Name -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll
                }
            
                # Process files in parallel
                $files = $rootItems | Where-Object { -not $_.Folder }
                if ($files.Count -gt 0) {
                    $files | ForEach-Object -Parallel {
                        $item = $_
                        $itemPath = Join-Path -Path $using:ExportFolder -ChildPath $item.Name
                    
                        # Using script block to call the function in the parent scope
                        & $using:SaveFileScriptBlock -DriveItemId $item.Id -DestinationPath $itemPath -Overwrite:$using:Overwrite -OverwriteAll:$using:OverwriteAll
                    } -ThrottleLimit 5
                }
            }
            else {
                Write-Host "Using PowerShell 5.1 sequential processing"
            
                # Process each item sequentially
                foreach ($item in $rootItems) {
                    $itemPath = Join-Path -Path $ExportFolder -ChildPath $item.Name
                
                    if ($item.Folder) {
                        # Process folder
                        Save-OneDriveFolder -DriveItemId $item.Id -DestinationPath $ExportFolder -FolderName $item.Name -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll
                    }
                    else {
                        # Download file
                        Save-OneDriveFile -DriveItemId $item.Id -DestinationPath $itemPath -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll
                    }
                }
            }
        
            # Complete the progress bar
            Write-Progress -Activity "Backing up OneDrive" -Completed
        
            # Report final statistics
            $elapsedTime = (Get-Date) - $Script:StartTime
            $elapsedText = "{0:hh\:mm\:ss}" -f $elapsedTime
            Write-Host "OneDrive backup completed in $elapsedText."
            Write-Host "Summary:"
            Write-Host "  - Total files: $Script:TotalFiles"
            Write-Host "  - Files processed: $Script:ProcessedFiles"
            Write-Host "  - Files skipped: $Script:SkippedFiles"
            Write-Host "  - Errors: $Script:ErrorCount"
            Write-Host "  - Output location: $ExportFolder"
        
            return $true
        }
        catch {
            Write-Progress -Activity "Backing up OneDrive" -Completed
            Write-Error -Exception ([System.Management.Automation.PSInvalidOperationException]::new(
                "Backup failed: $_",
                $_.Exception
            ))
            return $false
        }
    }

    # Main script execution
    try {
        # Initialize modules
        if (-not (Initialize-GraphModules -UpdateModules:$UpdateModules)) {
            throw "Failed to initialize required modules"
        }
    
        # Connect to Microsoft Graph
        if (-not (Connect-OneDriveGraph)) {
            throw "Failed to connect to Microsoft Graph"
        }
    
        # Start backup
        if (-not (Backup-OneDrivePersonal -ExportFolder $ExportFolder -Overwrite:$Overwrite -OverwriteAll:$OverwriteAll)) {
            throw "Backup operation failed"
        }
    }
    catch {
        Write-Error "Script execution failed: $_"
        exit 1
    }
}