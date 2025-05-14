function Get-FolderSizes {
<#
.SYNOPSIS
    Generates an interactive HTML report of folder sizes with sorting and filtering capabilities.

.DESCRIPTION
    This script analyzes folder sizes in a specified path and creates a detailed HTML report with the following features:
    - Interactive sorting by folder name, size, and last modified date
    - Search/filter functionality for folders
    - Responsive design with alternating row colors
    - Size formatting in appropriate units (B, KB, MB, GB, TB)
    - Progress tracking for long operations
    - Configurable minimum size filter and folder depth
    - Summary statistics including total size and folder count

.PARAMETER StartingPath
    The root path to start analyzing. Default is "C:\Temp\"

.PARAMETER OutputPath
    The path where the HTML report will be saved. Default is "C:\Temp\"

.PARAMETER SortBy
    Initial sort column. Options: 'Folder Name','Last Change','Size'. Default is "Size"

.PARAMETER MinSize
    Minimum file size in bytes to include in report. Default is 102400 (100KB)

.PARAMETER Depth
    Maximum folder depth to analyze. Must be greater than StartingPath depth. Default is 3

.PARAMETER IncludeZero
    Switch to include empty folders and those smaller than MinSize

.PARAMETER ShowResults
    Switch to automatically open the report in default browser when complete

.EXAMPLE
    Get-FolderSizes -StartingPath "C:\Users\myUser" -OutputPath "C:\Temp\Report" -Depth 5

.NOTES
    Requires PowerShell 5.1 or later
    Author: differentstuff
    Last Modified: 24.01.2025
#>
[CmdletBinding()]
Param (
    [Parameter(ValueFromPipeline)]
	[string]$StartingPath = "C:\Temp\",
	[string]$OutputPath = "C:\Temp\",
    [string]$SortBy = "Size", # 'Folder Name','Last Change','Size'
    [int]$MinSize = 102400, # exclude files smaller than 100KB (102400 B)
    [int]$Depth = 3, # Depth of Path to show in Report (must be greater than StartingPath)
    [switch]$IncludeZero, # Includes empty and small files/folders under 0.01MB
    [switch]$ShowResults
)

Function Set-AlternatingRows {
    [CmdletBinding()]
       Param(
           [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
        [object[]]$Lines,
       
           [Parameter(Mandatory=$True)]
           [string]$CSSEvenClass,
       
        [Parameter(Mandatory=$True)]
           [string]$CSSOddClass
       )
    Begin {
        $ClassName = $CSSEvenClass
    }
    Process {
        ForEach ($Line in $Lines)
        {	$Line = $Line.Replace("<tr>","<tr class=""$ClassName"">")
            If ($ClassName -eq $CSSEvenClass)
            {	$ClassName = $CSSOddClass
            }
            Else
            {	$ClassName = $CSSEvenClass
            }
            Return $Line
        }
    }
}

function Get-FolderSize {
    param (
        [string]$Path
        )
    try {
        $totalSize = (Get-ChildItem -LiteralPath $Path -File -Recurse -Force -ErrorAction SilentlyContinue -ErrorVariable accessErrors | 
            Measure-Object -Property Length -Sum).Sum
        if ($accessErrors) {
            Write-Warning "Some files in $Path could not be accessed"
        }

        if($totalSize){
            return $totalSize
        }
        else{
            return 0
        }
    }
    catch {
        Write-Warning "Error accessing $Path : $_"
        return 0
    }
}

function Get-FolderReport {
    param (
        [string]$RootPath,
        [int]$MaxDepth,
        [switch]$IncludeZero,
        [int]$MinSize,
        [ref]$TotalProcessedDirs
    )
    
    Write-Verbose "Starting folder report generation from root: $RootPath"
    Write-Verbose "Parameters - MaxDepth: $MaxDepth, MinSize: $MinSize"

    $results = [System.Collections.ArrayList]::new()
    $processedPaths = @{}
    
    # Get all folders first for progress reporting
    $allFolders = @($RootPath) # Add Base folder
    $startingDepth = ($RootPath.Split([IO.Path]::DirectorySeparatorChar) | Where-Object { $_ }).Count
    Write-Verbose "Starting with depth: $startingDepth"

    if ($startingDepth -gt $MaxDepth) {
        Write-Warning "Starting folder depth ($startingDepth) is greater than specified MaxDepth ($MaxDepth). No results will be shown."
        return $results  # Return empty results since nothing will match the depth criteria
    }

    Write-Verbose "Processing recursively (This may take a while)"
    
    # Only get folders up to MaxDepth
    $remainingDepth = $MaxDepth - $startingDepth
    Write-Verbose "Remaining depth to process: $remainingDepth"
    if ($remainingDepth -gt 0) {
        $allFolders += @(Get-ChildItem -LiteralPath $RootPath -Directory -Recurse -ErrorAction SilentlyContinue | 
            Where-Object { ($_.FullName.Split([IO.Path]::DirectorySeparatorChar).Count) -le $MaxDepth } |
            Select-Object -ExpandProperty FullName)
    }
    
    $totalFolders = $allFolders.Count
    Write-Verbose "Total folders to process: $totalFolders"
    $processedCount = 0
    
    foreach ($folder in $allFolders) {
        $processedCount++
        Write-Progress -Activity "Processing Folders" `
            -Status "Processing ($processedCount of $totalFolders): $folder" `
            -PercentComplete (($processedCount / $totalFolders) * 100)
        
        if ($processedPaths.ContainsKey($folder)) { continue }

        try {
            # Check MaxDepth first
            $folderDepth = ($folder.Split([IO.Path]::DirectorySeparatorChar).Count)
            if ($folderDepth -gt $MaxDepth) {
                # Write-Verbose "Skipping: $folder - exceeds maximum depth"
                continue  # Skip to next folder
            }
            else{
                Write-Verbose "Processing folder: $folder (Depth: $folderDepth / Max: $MaxDepth)"
            }

            # Only process folders within depth limit
            $folderSize = Get-FolderSize -Path $folder

            if ($folderSize -ge $MinSize -or $IncludeZero) {
                # Only get folder info when actually needed
                Write-Verbose " + Adding folder to results (Size: $([Math]::round(($folderSize/1MB),2)) MB)"
                $folderInfo = Get-Item -LiteralPath $folder -Force -ErrorAction Stop
                $null = $results.Add([PSCustomObject]@{
                    'Folder Name' = $folderInfo.FullName
                    'Size' = [long]$folderSize
                    'Last Change' = $folderInfo.LastWriteTime
                })
            }
            else {
                Write-Verbose " - Skipping folder in results (Size: $([Math]::round(($folderSize/1MB),2)) MB)"
            }
        
            $TotalProcessedDirs.Value++
            $processedPaths[$folder] = $true
        }
        catch {
            Write-Warning "Error processing $folder : $_"
        }
    }
    
    Write-Progress -Activity "Processing Folders" -Completed
    Write-Verbose "Folder report generation complete. Total folders processed: $($TotalProcessedDirs.Value)"
    return $results
}

function Get-HTMLStyle {
    param(
        [string]$Path,
        $TotalProcessedDirs,
        $TotalProcessedSize
    )

    function Format-ByteSize {
        param ([long]$Size)
        if ($Size -eq 0) { return "0 B" }
        $sizes = 'B','KB','MB','GB','TB'
        $index = [Math]::Floor([Math]::Log($Size, 1024))
        $value = [Math]::Round($Size / [Math]::Pow(1024, $index), 2)
        return "$value $($sizes[$index])"
    }

$Header = @"
<meta charset="UTF-8">
<title>Folder Sizes for "$Path"</title>
<style>
:root {
    --primary-color: #6495ED;
    --primary-hover: #4F7BE3;
    --border-color: #ddd;
    --hover-bg: #f5f5f5;
    --even-row: #dddddd;
    --odd-row: #ffffff;
}

* {
    font-family: Arial, Helvetica, sans-serif;
    box-sizing: border-box;
    margin: 0;
    padding: 0;
}

body {
    padding: 20px;
    max-width: 1400px;
    margin: 0 auto;
}

h1 {
    color: var(--primary-color);
    margin-bottom: 20px;
}

.container {
    box-shadow: 0 0 10px rgba(0,0,0,0.1);
    border-radius: 8px;
    overflow: hidden;
}

TABLE {
    width: 100%;
    border-collapse: collapse;
    margin: 0;
    background: white;
}

TH {
    padding: 12px 8px;
    background-color: var(--primary-color);
    color: white;
    text-align: left;
    position: sticky;
    top: 0;
    cursor: pointer;
    user-select: none;
    transition: background-color 0.2s;
    z-index: 10;
}

TH:after {
    content: '\21F3';
    color: rgba(255, 255, 255, 0.5);
    font-size: 0.8em;
    margin-left: 5px;
}

TH.asc:after {
    content: '\21E7';
    color: white;
}

TH.desc:after {
    content: '\21E9';
    color: white;
}

TH:hover {
    background-color: var(--primary-hover);
}

TD {
    padding: 12px 8px;
    border-bottom: 1px solid var(--border-color);
    max-width: 0;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    transition: all 0.2s;
}

TD:hover {
    white-space: normal;
    word-break: break-all;
    background-color: var(--hover-bg);
}

.odd  { background-color: var(--odd-row); }
.even { background-color: var(--even-row); }

TR:hover TD {
    background-color: var(--hover-bg);
}

TD:nth-child(1), TH:nth-child(1) { width: 60%; }
TD:nth-child(2), TH:nth-child(2) { width: 20%; text-align: right; }
TD:nth-child(3), TH:nth-child(3) { width: 20%; }

.report-footer {
    margin-top: 20px;
    display: flex;
    justify-content: center;
}

.summary-container {
    max-width: 600px;
    background: #f8f9fa;
    border-radius: 6px;
    padding: 15px 25px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
}

.summary-stats {
    display: flex;
    gap: 30px;
    justify-content: center;
    margin-bottom: 10px;
}

.stats {
    margin: 20px 0;
    padding: 15px;
    background: #f8f9fa;
    border-radius: 4px;
    display: flex;
    justify-content: space-between;
    flex-wrap: wrap;
}

.stat-item {
    display: flex;
    padding: 10px;
    gap: 10px;
    align-items: center;
}

.stat-label {
    color: #666;
    font-size: 0.9em;
}

.stat-value {
    color: var(--primary-color);
    font-size: 1.2em;
    font-weight: bold;
    font-family: monospace;
}

.timestamp {
    text-align: center;
    color: #888;
    font-size: 0.8em;
    margin-top: 8px;
    padding-top: 8px;
    border-top: 1px solid #eee;
}

.params-info {
    text-align: center;
    color: #666;
    font-size: 0.85em;
    margin-top: 8px;
    padding: 8px 0;
    border-top: 1px solid #eee;
}

.param-item {
    display: inline-block;
    margin: 0 10px;
    padding: 2px 8px;
    background: #eee;
    border-radius: 4px;
    font-family: monospace;
}

@media (max-width: 768px) {
    .stat-block {
        flex-direction: column;
        align-items: flex-start;
        gap: 5px;
    }
    TD:nth-child(3), TH:nth-child(3) { display: none; }
    TD:nth-child(1), TH:nth-child(1) { width: 70%; }
    TD:nth-child(2), TH:nth-child(2) { width: 30%; }
}

#searchInput {
    width: 100%;
    padding: 12px 20px;
    margin: 8px 0;
    border: 2px solid var(--primary-color);
    border-radius: 4px;
    font-size: 16px;
    transition: border-color 0.3s;
}

#searchInput:focus {
    outline: none;
    border-color: var(--primary-hover);
    box-shadow: 0 0 5px rgba(100,149,237,0.3);
}

</style>
<script>
let lastSortedColumn = 1; // Default to size column
let lastSortDirection = 'desc'; // Default to descending

function formatBytes(bytes) {
    if (bytes === 0) return '0 KB';
    const units = ['B', 'KB', 'MB', 'GB', 'TB'];
    const k = 1024;
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + units[i];
}
    
function parseCustomDate(dateStr) {
    // Parse date in format "dd.MM.yyyy HH:mm:ss"
    const [day, month, year, hours, minutes, seconds] = dateStr.match(/\d+/g).map(Number);
    return new Date(year, month - 1, day, hours, minutes, seconds); // Month is 0-based
}

function sortTable(n) {
    console.log('Sorting column:', n); // Debug output
    const table = document.getElementsByTagName("TABLE")[0];
    const headers = Array.from(table.getElementsByTagName("TH"));
    const rows = Array.from(table.rows).slice(1);
    
    // Update sort direction
    if (n === lastSortedColumn) {
        lastSortDirection = lastSortDirection === 'asc' ? 'desc' : 'asc';
    } else {
        lastSortDirection = 'asc';
    }
    
    // Update header styles
    headers.forEach(header => {
        header.style.backgroundColor = 'var(--primary-color)';
        header.classList.remove('asc', 'desc');
    });
    headers[n].style.backgroundColor = 'var(--primary-hover)';
    headers[n].classList.add(lastSortDirection);

    // Sort function
    const comparator = (a, b) => {
        let x, y;
        if (n === 1) { // Size column
            x = parseInt(a.cells[n].getAttribute('data-raw-size')) || 0;
            y = parseInt(b.cells[n].getAttribute('data-raw-size')) || 0;
        } else if (n === 2) { // Date column
            x = parseCustomDate(a.cells[n].textContent.trim());
            y = parseCustomDate(b.cells[n].textContent.trim());
        } else { // Text column (folder name)
            x = a.cells[n].textContent.trim();
            y = b.cells[n].textContent.trim();
            return lastSortDirection === 'asc' ? 
                x.localeCompare(y, undefined, {numeric: true, sensitivity: 'base'}) :
                y.localeCompare(x, undefined, {numeric: true, sensitivity: 'base'});
        }
        
        return lastSortDirection === 'asc' ? x - y : y - x;
    };
    
    // Sort and reinsert rows
    rows.sort(comparator).forEach(row => table.tBodies[0].appendChild(row));
    lastSortedColumn = n;
    
    // Update row classes
    updateRowClasses();
}

function updateRowClasses() {
    const rows = document.getElementsByTagName('tr');
    for (let i = 1; i < rows.length; i++) {
        rows[i].classList.remove('odd', 'even');
        rows[i].classList.add(i % 2 ? 'odd' : 'even');
    }
}

function filterTable() {
    const input = document.getElementById("searchInput");
    const filter = input.value.toLowerCase();
    const table = document.getElementsByTagName("table")[0];
    const rows = table.getElementsByTagName("tr");
    let visibleCount = 0;

    for (let i = 1; i < rows.length; i++) {
        const td = rows[i].getElementsByTagName("td")[0];
        if (td) {
            const txtValue = td.textContent || td.innerText;
            const isVisible = txtValue.toLowerCase().indexOf(filter) > -1;
            rows[i].style.display = isVisible ? "" : "none";
            if (isVisible) visibleCount++;
        }
    }

    // Update stats
    document.getElementById('visible-count').textContent = visibleCount;
    updateRowClasses();
}

// Initialize sorting indicators on load
window.onload = function() {
    console.log('Window loaded'); // Debug output
    const headers = document.getElementsByTagName("TH");
    headers[1].style.backgroundColor = 'var(--primary-hover)';
    headers[1].classList.add('desc');
    
    // Format sizes
    formatTableSizes();

    // Format total size in footer
    const totalSizeElement = document.querySelector('.summary-stats .stat-value[data-raw-size]');
    if (totalSizeElement) {
        const rawSize = parseInt(totalSizeElement.getAttribute('data-raw-size'));
        totalSizeElement.textContent = formatBytes(rawSize);
    }

    // Calculate initial statistics
    const table = document.getElementsByTagName("table")[0];
    document.getElementById('total-count').textContent = table.rows.length - 1;
    document.getElementById('visible-count').textContent = table.rows.length - 1;
};

function formatTableSizes() {
    const table = document.getElementsByTagName("table")[0];
    const rows = table.getElementsByTagName("tr");
    
    // Start from 1 to skip header row
    for (let i = 1; i < rows.length; i++) {
        const sizeCell = rows[i].cells[1];
        const rawSize = parseInt(sizeCell.textContent);
        // Store raw value as data attribute for sorting
        sizeCell.setAttribute('data-raw-size', rawSize);
        // Display formatted size
        sizeCell.textContent = formatBytes(rawSize);
    }
}

</script>
"@

$Pre = @"
<h1>Folder Sizes Report for: $Path</h1>
<div class="stats">
    <div class="stat-item">
        <div>Total Folders</div>
        <div class="stat-value" id="total-count">0</div>
    </div>
    <div class="stat-item">
        <div>Visible Folders</div>
        <div class="stat-value" id="visible-count">0</div>
    </div>
</div>
<input type="text" id="searchInput" onkeyup="filterTable()" placeholder="Search for folders...">
<div class="container">
"@

$Post = @"
<div class="report-footer">
    <div class="summary-container">
        <div class="summary-stats">
            <div class="stat-item">
                <span class="stat-label">Total Folders processed:</span>
                <span class="stat-value">$TotalProcessedDirs</span>
            </div>
            <div class="stat-item">
                <span class="stat-label">Total Size:</span>
                <span class="stat-value" data-raw-size="$TotalProcessedSize"></span>
            </div>
        </div>
        <div class="params-info">
            <span class="param-item">Depth: $(if($Depth -eq 0){"∞"}else{$Depth})</span>
            <span class="param-item">Min Size: $(Format-ByteSize $MinSize)</span>
        </div>
        <div class="timestamp">Generated on $(Get-Date -Format "dd.MM.yyyy HH:mm:ss")</div>
    </div>
</div>
</body></html>
"@
    
    return @{
        Header = $Header
        Pre = $Pre
        Post = $Post
    }
}

# Main execution
Write-Verbose "Script started at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Verbose "Parameters - StartingPath: $StartingPath, Depth: $Depth, MinSize: $MinSize"
$Report = [System.Collections.ArrayList]::new()
$TotalProcessedDirs = 0
$processedPaths = @{}

if ($(($StartingPath.Split([IO.Path]::DirectorySeparatorChar) | Where-Object { $_ }).Count) -gt $Depth) {
    throw "Starting folder depth ($StartingPath) is greater than specified maximum depth ($Depth). No results will be shown."
}

# Count total folders across all paths
if (Test-Path $StartingPath) {
    Write-Verbose "Gathering Folder Information"

    Write-Verbose "Getting Total Folder Size (This may take a while)"
    $TotalProcessedSize = Get-FolderSize -Path $StartingPath

    Write-Verbose "Getting Results"
    $results = Get-FolderReport -RootPath $StartingPath `
                               -MaxDepth $Depth `
                               -IncludeZero:$IncludeZero `
                               -MinSize $MinSize `
                               -TotalProcessedDirs ([ref]$TotalProcessedDirs) `

    if ($results -is [System.Collections.ICollection]) {
        $Report.AddRange($results)
    } else {
        $null = $Report.Add($results)
    }
} else {
    Write-Warning "Path not found: $StartingPath"
}

# Clear the progress bar when done
Write-Progress -Activity "Processing Folders" -Completed -Id 0

Write-Verbose "Starting HTML report generation with $($Report.Count) folders"
# Create new objects with Select-Object
$ReportSorted = $Report | Select-Object * | Sort-Object $SortBy -Descending
if(!$IncludeZero){
    # Remove small values
    $ReportSorted = $ReportSorted | Where-Object {$_.Size -gt $MinSize}
}

#Create the HTML for our report
$HTMLStyle = Get-HTMLStyle -Path $StartingPath -TotalProcessedDirs $TotalProcessedDirs -TotalProcessedSize $TotalProcessedSize

#Create the report and save it to a file
Write-Verbose "Sorting results by $SortBy"
$HTML = $ReportSorted | 
Sort-Object { $_.Size } -Descending |
Select-Object @{
    Name='Folder Name'; 
    Expression={ $_.'Folder Name' }
},
@{
    Name='Size'; 
    Expression={ $_.Size }
},
@{
    Name='Last Change'; 
    Expression={ $_.'Last Change' }
} |
ConvertTo-Html -PreContent $HTMLStyle.Pre -PostContent $HTMLStyle.Post -Head $HTMLStyle.Header |
Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd

# Add onclick handlers to the headers
$HTML = $HTML -replace '<th>Folder Name</th>', '<th onclick="sortTable(0)">Folder Name</th>'
$HTML = $HTML -replace '<th>Size</th>', '<th onclick="sortTable(1)">Size</th>'
$HTML = $HTML -replace '<th>Last Change</th>', '<th onclick="sortTable(2)">Last Change</th>'

# Construct export name (e.g.: C:\Temp\FolderSizes-Temp-23012025-1226.html)
$HTMLPath = Join-Path $OutputPath "FolderSizes-$(Split-Path $StartingPath -Leaf)-$(Get-Date -Format "ddMMyyyy-HHmm").html"
Write-Verbose "Saving HTML report"

# Export file
if(!(Get-item ([system.io.fileinfo]$HTMLPath).DirectoryName)){
        Write-Host -ForegroundColor Yellow "Creating Folder for: $HTMLPath"
        New-Item $HTMLPath -ItemType Directory
    }
Out-File -InputObject $HTML -FilePath $HTMLPath
Write-Verbose "Report saved to Path: $HTMLPath"

#Display the report in your default browser
if($ShowResults){
    Write-Verbose "$(Get-Date): Opening Results"
    & $HTMLPath
}

Write-Verbose "Script completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Verbose "Total folders processed: $TotalProcessedDirs"
Write-Verbose "Total size processed: $([Math]::round(($TotalProcessedSize/1MB),2)) MB"
}
