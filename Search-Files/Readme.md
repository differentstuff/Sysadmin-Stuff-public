# Search-Files PowerShell Tool

A high-performance PowerShell script for searching files across large directory structures with advanced filtering capabilities.

## Features

- **Fast Directory Enumeration**: Uses parallel processing (PS7+) or optimized sequential processing (PS5)
- **Date Range Filtering**: Search files by last modified date
- **Path Exclusions**: Skip specific directories during search
- **File Pattern Matching**: Support for wildcards (*.prt, *.txt, etc.)
- **Owner Information**: Attempts to capture last modified user when possible
- **Crash Recovery**: Results are written immediately to CSV for recovery
- **Progress Tracking**: Real-time progress updates during execution

## Requirements

- PowerShell 5.0+ (PowerShell 7+ recommended for parallel processing)
- Read access to target directories
- Write access to output location

## Usage

### Basic Usage
```powershell
Search-Files -RootPath "E:\Data"
```

### Advanced Usage
```powershell
# Option 1
Search-Files -RootPath "D:\ProjectFiles" `
             -FileFilter "*.dwg" `
             -StartDate "2025-11-01" `
             -EndDate "2025-11-30" `
             -ExcludePaths @("D:\ProjectFiles\Archive", "D:\ProjectFiles\Temp") `
             -OutputFile "CADFiles.csv"

# Option 2
$excludePaths = @(
	"E:\Data\1",
	"E:\Data\2",
	"E:\Data\3",
	"E:\Data\4"
)
Search-Files -RootPath E:\Data\ -ExcludePaths $excludePaths

```

## Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `RootPath` | String | *Required* | Root directory to start searching |
| `FileFilter` | String | `"*.prt"` | File pattern to match (wildcards supported) |
| `ExcludePaths` | String[] | `@()` | Array of paths to exclude (startswith match) |
| `StartDate` | DateTime | `2024-10-01` | Start of date range filter |
| `EndDate` | DateTime | `2024-10-20` | End of date range filter |
| `OutputFile` | String | `"FileSearchResults.csv"` | Output CSV file path |

## Output Format

The tool generates a CSV file with the following columns:
- `filename`: Name of the file
- `full_path`: Complete file path
- `last_modified`: Last modification timestamp (yyyy-MM-dd HH:mm:ss)
- `last_modified_user`: File owner (when accessible)

## Performance Notes

- **PowerShell 7+**: Uses parallel processing for faster execution on large directory structures
- **PowerShell 5**: Falls back to sequential processing but still optimized
- **Error Handling**: Silently skips files/directories with access restrictions
- **Memory Efficient**: Processes files incrementally to handle large datasets

## Examples

### Search for CAD Files
```powershell
Search-Files -RootPath "\\server\cad" -FileFilter "*.dwg"
```

### Search with Date Range
```powershell
Search-Files -RootPath "C:\Projects" -StartDate "2024-01-01" -EndDate "2024-12-31"
```

### Exclude Temporary Directories
```powershell
Search-Files -RootPath "D:\Data" -ExcludePaths @("D:\Data\temp", "D:\Data\backup")
```

## Troubleshooting

- **Access Denied**: The script silently skips inaccessible files/folders
- **Large Datasets**: For 100k+ directories, expect longer enumeration times
- **Network Drives**: Performance may vary; consider running directly on the server
- **Memory Usage**: Monitor system resources with very large directory structures
