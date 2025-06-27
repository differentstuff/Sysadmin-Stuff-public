function Search-Files {
<#
.SYNOPSIS
    Searches for files in a directory structure with various filtering options.
    HDD-optimized file search with sequential directory traversal.
.DESCRIPTION
    This script searches for files matching specified criteria, including date range,
    file pattern, and exclusion paths. Results are saved to a CSV file.
.PARAMETER RootPath
    The root directory to start searching from.
.PARAMETER FileFilter
    The file pattern to search for (default: *.prt).
.PARAMETER ExcludePaths
    Array of paths to exclude (startswith match).
.PARAMETER StartDate
    Start of last modified date range (default: 2024-10-01).
.PARAMETER EndDate
    End of last modified date range (default: 2024-10-30).
.PARAMETER OutputFile
    Output CSV file path (default: FileSearchResults.csv in current directory).
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$RootPath,
    
    [string]$FileFilter = "*.prt",
    
    [string[]]$ExcludePaths = @(),
    
    [datetime]$StartDate = "2024-10-01",
    
    [datetime]$EndDate = "2024-10-30",
    
    [string]$OutputFile = "FileSearchResults.csv",
    
    [switch]$UseSSD = $false  # Set to $true if running on SSD for parallel processing
)

#region Functions

# HDD-optimized recursive directory traversal (depth-first, single-threaded)
function Search-DirectoryRecursive {
    param (
        [string]$Path,
        [string]$Filter,
        [datetime]$Start,
        [datetime]$End,
        [string[]]$ExcludePaths,
        [string]$OutputPath,
        [ref]$ProcessedCount
    )
    
    # Check if current path should be excluded
    foreach ($exclude in $ExcludePaths) {
        if ($Path.StartsWith($exclude, [System.StringComparison]::OrdinalIgnoreCase)) {
            return
        }
    }
    
    try {
        # Process files in current directory first (sequential disk access)
        $files = [System.IO.Directory]::EnumerateFiles($Path, $Filter)
        foreach ($filePath in $files) {
            try {
                $fileInfo = [System.IO.FileInfo]::new($filePath)
                if ($fileInfo.LastWriteTime -ge $Start -and $fileInfo.LastWriteTime -le $End) {
                    $lastModifiedUser = $null
                    try {
                        # Only get ACL if really needed (expensive operation)
                        $acl = Get-Acl -Path $fileInfo.FullName -ErrorAction SilentlyContinue
                        $lastModifiedUser = $acl.Owner
                    } catch {
                        # Silently continue
                    }
                    
                    $csvLine = '"{0}","{1}","{2}","{3}"' -f 
                        $fileInfo.Name,
                        $fileInfo.FullName,
                        $fileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"),
                        $lastModifiedUser
                    
                    # Append immediately for crash recovery
                    $csvLine | Out-File -FilePath $OutputPath -Append -Encoding UTF8
                }
            } catch {
                # Silently continue on file errors
            }
        }
        
        # Then process subdirectories (depth-first keeps disk head in same area)
        $subDirs = [System.IO.Directory]::EnumerateDirectories($Path)
        foreach ($subDir in $subDirs) {
            $shouldExclude = $false
            foreach ($exclude in $ExcludePaths) {
                if ($subDir.StartsWith($exclude, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $shouldExclude = $true
                    break
                }
            }
            
            if (-not $shouldExclude) {
                Search-DirectoryRecursive -Path $subDir -Filter $Filter -Start $Start -End $End -ExcludePaths $ExcludePaths -OutputPath $OutputPath -ProcessedCount $ProcessedCount
            }
        }
        
        # Update progress every 50 directories to reduce overhead
        $ProcessedCount.Value++
        if ($ProcessedCount.Value % 50 -eq 0) {
            Write-Host "Processed $($ProcessedCount.Value) directories..." -NoNewline
            Write-Host "`r" -NoNewline
        }
        
    } 
    catch {
        # Silently continue on directory access errors
    }
}

# SSD-optimized parallel version
function Search-DirectoriesParallel {
    param (
        [string[]]$Directories,
        [string]$Filter,
        [datetime]$Start,
        [datetime]$End,
        [string]$OutputPath
    )
    
    $Directories | ForEach-Object -Parallel {
        $dir = $_
        try {
            $files = [System.IO.Directory]::EnumerateFiles($dir, $using:Filter)
            foreach ($filePath in $files) {
                try {
                    $fileInfo = [System.IO.FileInfo]::new($filePath)
                    if ($fileInfo.LastWriteTime -ge $using:Start -and $fileInfo.LastWriteTime -le $using:End) {
                        $lastModifiedUser = $null
                        try {
                            $acl = Get-Acl -Path $fileInfo.FullName -ErrorAction SilentlyContinue
                            $lastModifiedUser = $acl.Owner
                        } catch {
                        }
                        
                        $csvLine = '"{0}","{1}","{2}","{3}"' -f 
                            $fileInfo.Name,
                            $fileInfo.FullName,
                            $fileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"),
                            $lastModifiedUser
                        
                        $csvLine | Out-File -FilePath $using:OutputPath -Append -Encoding UTF8
                    }
                } catch {
                }
            }
        } catch {
        }
    } -ThrottleLimit 4  # Lower throttle for SSD
}

#endregion Functions

#region Main
# Initialize CSV file
$csvHeaders = "filename,full_path,last_modified,last_modified_user"
$csvHeaders | Out-File -FilePath $OutputFile -Force -Encoding UTF8

Write-Host "Starting file search..."
Write-Host "Root: $RootPath"
Write-Host "Filter: $FileFilter"
Write-Host "Date Range: $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))"
Write-Host "Storage Type: $(if($UseSSD){'SSD (Parallel)'}else{'HDD (Sequential)'})"

$processedCount = 0
$startTime = Get-Date

if ($UseSSD -and $PSVersionTable.PSVersion.Major -ge 7) {
    # SSD: Use parallel processing
    Write-Host "Enumerating directories for parallel processing..."
    try {
        $allDirs = Get-ChildItem -Path $RootPath -Directory -Recurse -ErrorAction SilentlyContinue | 
            Where-Object {
                $dirPath = $_.FullName
                $include = $true
                foreach ($exclude in $ExcludePaths) {
                    if ($dirPath.StartsWith($exclude, [System.StringComparison]::OrdinalIgnoreCase)) {
                        $include = $false
                        break
                    }
                }
                $include
            } | ForEach-Object { $_.FullName }
        
        $allDirs = @($RootPath) + $allDirs
        Write-Host "Processing $($allDirs.Count) directories in parallel..."
        Search-DirectoriesParallel -Directories $allDirs -Filter $FileFilter -Start $StartDate -End $EndDate -OutputPath $OutputFile
    } 
    catch {
        Write-Error "Error in parallel processing: $_"
    }
} 
else {
    # HDD: Use sequential recursive traversal
    Write-Host "Using HDD-optimized sequential traversal..."
    $processedCountRef = [ref]$processedCount
    Search-DirectoryRecursive -Path $RootPath -Filter $FileFilter -Start $StartDate -End $EndDate -ExcludePaths $ExcludePaths -OutputPath $OutputFile -ProcessedCount $processedCountRef
}

$endTime = Get-Date
$duration = $endTime - $startTime

Write-Host ""
Write-Host "Search completed in $($duration.TotalMinutes.ToString('F1')) minutes"
Write-Host "Results saved to: $OutputFile"

# Show file count
try {
    $resultCount = (Get-Content $OutputFile | Measure-Object).Count - 1  # -1 for header
    Write-Host "Files found: $resultCount"
} 
catch {
    Write-Host "Could not count results file"
}

#endregion Main
}