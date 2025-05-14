function Get-DirectoryTree {
<#
.SYNOPSIS
Displays a directory tree for the specified path.

Author: Jean-Marie Heck (hecjex)
License: BSD 3-Clause
Required Dependencies: None

.DESCRIPTION
This function displays a directory tree for the specified path, including subdirectories and optionally files.

.PARAMETER Path
The path to the directory to display the tree for. This parameter is mandatory and must be a valid directory path.

.PARAMETER IncludeFiles
A switch parameter that specifies whether to include files in the directory tree display. By default, only directories are displayed.

.PARAMETER IncludeSize
A switch parameter that adds file sizes in MB to the display. Only applies to files, not directories.

.PARAMETER IndentSize
The number of spaces to use for indentation in the directory tree display. The default is 2.

.PARAMETER MaxDepth
The maximum depth of the directory tree to display. The default is unlimited.

.PARAMETER CleanOutput
A switch parameter that removes all symbols and only shows indentation.

.EXAMPLE
Get-DirectoryTree -Path "C:\Users"

.EXAMPLE
Get-DirectoryTree -Path "C:\Users" -IncludeFiles -IncludeSize

.EXAMPLE
Get-DirectoryTree -Path "C:\Users" -IndentSize 2 -MaxDepth 3 -CleanOutput

.INPUTS
String
Accepts a string representing the path to the directory to display the tree for.

.OUTPUTS
None
This function does not output any objects, but instead displays the directory tree directly to the console.

.LINK
https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/get-childitem
https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/resolve-path
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false,
                   Position = 0,
                   ValueFromPipeline = $true)]
        [ValidateScript({Test-Path $_})]
        [string]$Path = $pwd,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeFiles,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeSize,

        [Parameter(Mandatory = $false)]
        [int]$IndentSize = 2,

        [Parameter(Mandatory = $false)]
        [int]$MaxDepth = [int]::MaxValue,

        [Parameter(Mandatory = $false)]
        [switch]$CleanOutput
    )

    $TreeSymbols = @{
        Directory = "+"    
        File      = "."    
        Space     = " "    
    }

    function Format-FileSize {
        param ([int64]$Size)
        try {
            return [string]::Format("[{0:N2} MB]", $Size / 1MB)
        }
        catch {
            return "[?.?? MB]"
        }
    }

    function Write-TreeNode {
        param(
            [string]$Path,
            [int]$CurrentDepth = 0,
            [string]$Indent = ""
        )

        if ($CurrentDepth -gt $MaxDepth) { return }

        $nodeName = Split-Path -Leaf $Path
        $isDirectory = Test-Path -Path $Path -PathType Container

        # Get size info for files if needed
        $sizeInfo = ""
        if ($IncludeSize -and -not $isDirectory -and -not $CleanOutput) {
            try {
                $size = (Get-Item -LiteralPath $Path -ErrorAction Stop).Length
                $sizeInfo = " $(Format-FileSize $size)"
            }
            catch {
                $sizeInfo = " [?.?? MB]"
            }
        }

        # Determine symbol
        $symbol = if ($CleanOutput) {
            ""
        } else {
            if ($isDirectory) { $TreeSymbols.Directory } else { $TreeSymbols.File }
        }

        Write-Host "$Indent$symbol $nodeName$sizeInfo"

        if ($isDirectory) {
            try {
                $allItems = Get-ChildItem -LiteralPath $Path -ErrorAction Stop
                $items = @()
                $items += $allItems | Where-Object { $_.PSIsContainer }
                if ($IncludeFiles) {
                    $items += $allItems | Where-Object { -not $_.PSIsContainer }
                }
                $items = $items | Sort-Object Name

                foreach ($item in $items) {
                    Write-TreeNode -Path $item.FullName `
                                 -CurrentDepth ($CurrentDepth + 1) `
                                 -Indent "$Indent$($TreeSymbols.Space * $IndentSize)"
                }
            }
            catch {
                Write-Warning "Access denied or error reading directory: $Path"
                return
            }
        }
    }

    try {
        $resolvedPath = (Resolve-Path $Path).Path
        if (-not $CleanOutput) {
            Write-Host "`nDirectory tree for: $resolvedPath`n"
        }
        Write-TreeNode -Path $resolvedPath
        if (-not $CleanOutput) {
            Write-Host ""
        }
    }
    catch {
        Write-Error "Error displaying tree for path $Path : $_"
    }
}
