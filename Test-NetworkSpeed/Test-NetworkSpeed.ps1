#Requires -Version 3.0
<#
.SYNOPSIS
    Measures network read/write throughput to a UNC path.
.DESCRIPTION
    Copies a set of local sample files to and from a remote UNC share,
    measures elapsed time, and returns a structured result object.
    Optionally generates random test files on the fly.
.PARAMETER LocalPath
    Local folder containing sample files for the test.
.PARAMETER RemotePath
    Target UNC path (e.g. \\server\share). A temporary 'SpeedTest' subfolder is created automatically and removed after use.
.PARAMETER Credential
    Optional PSCredential (single), or a [PSCustomObject] with .LocalCredentials and .RemoteCredentials
    PSCredential properties (as returned by -SetCredentials).
.PARAMETER GenerateFiles
    If set, generates random test files in LocalPath instead of using existing ones.
.PARAMETER FileCount
    Number of random files to generate. Default: 10.
.PARAMETER FileSizeMB
    Size of each generated file in MB. Default: 10.
.PARAMETER SetCredentials
    Interactive mode: prompts for two credentials and returns them as a structured object.
    No speed test is performed.
.OUTPUTS
    PSCustomObject with: Server, Timestamp, TotalSizeMB, WriteMbps, WriteTime, ReadMbps, ReadTime, Status
    OR (with -SetCredentials): PSCustomObject with .LocalCredentials and .RemoteCredentials PSCredential properties
.EXAMPLE
    Test-NetworkSpeed -LocalPath ".\files" -RemotePath "\\server1\share"
.EXAMPLE
    $creds = Test-NetworkSpeed -SetCredentials
    Test-NetworkSpeed -LocalPath ".\files" -RemotePath "\\server1\share" -Credential $creds
.EXAMPLE
    Test-NetworkSpeed -LocalPath ".\files" -RemotePath "\\server1\share" -GenerateFiles -FileCount 5 -FileSizeMB 20 -Verbose
#>
function Test-NetworkSpeed {
    [CmdletBinding(DefaultParameterSetName = 'Run')]
    [OutputType([PSCustomObject])]
    Param (
        [Parameter(Mandatory, Position = 0, ParameterSetName = 'Run')]
        [string]$LocalPath,

        [Parameter(Mandatory, Position = 1, ParameterSetName = 'Run')]
        [string]$RemotePath,

        [Parameter(ParameterSetName = 'Run')]
        [object]$Credential,

        [Parameter(ParameterSetName = 'Run')]
        [switch]$GenerateFiles,

        [Parameter(ParameterSetName = 'Run')]
        [ValidateRange(1, 1000)]
        [int]$FileCount = 10,

        [Parameter(ParameterSetName = 'Run')]
        [ValidateRange(1, 10240)]
        [int]$FileSizeMB = 10,

        [Parameter(Mandatory, ParameterSetName = 'SetCreds')]
        [switch]$SetCredentials
    )

    # -------------------------------------------------------------------------
    # CREDENTIAL HELPERS
    # -------------------------------------------------------------------------

    function Invoke-CredentialPrompt {
        param(
            [string]$Label,
            [string]$SkipKey   = '1',
            [string]$CopyKey   = $null,
            [string]$CopyLabel = $null
        )
        Write-Host ""
        Write-Host "  [$Label]" -ForegroundColor Cyan
        Write-Host "  Press $SkipKey to skip (no credentials)" -ForegroundColor DarkGray
        if ($CopyKey) {
            Write-Host "  Press $CopyKey to reuse $CopyLabel" -ForegroundColor DarkGray
        }
        Write-Host "  Any other key / Enter to enter credentials" -ForegroundColor DarkGray
        Write-Host ""

        # ReadKey fails when stdin is redirected (ISE, VS Code, piped). Fall back to Read-Host.
        if ([System.Console]::IsInputRedirected) {
            $key = (Read-Host "  Choice").Trim()
        } else {
            $key = [System.Console]::ReadKey($true).KeyChar.ToString()
        }

        if ($key -eq $SkipKey) {
            Write-Host "  Skipped" -ForegroundColor DarkGray
            return [System.Management.Automation.PSCredential]::Empty
        }
        if ($CopyKey -and $key -eq $CopyKey) {
            Write-Host "  Reusing $CopyLabel" -ForegroundColor DarkGray
            return $null  # caller replaces with first cred
        }

        return Get-Credential -Message "Enter credentials for $Label"
    }

    function Resolve-Credentials {
        param([object]$Cred)
        # Returns @{ LocalCredentials = <PSCredential>; RemoteCredentials = <PSCredential> }
        $empty = [System.Management.Automation.PSCredential]::Empty

        if ($null -eq $Cred) {
            return @{ LocalCredentials = $empty; RemoteCredentials = $empty }
        }

        # Dual-credential object (.LocalCredentials / .RemoteCredentials)
        if ($Cred.PSObject.Properties['LocalCredentials'] -and $Cred.PSObject.Properties['RemoteCredentials']) {
            return @{
                LocalCredentials  = if ($Cred.LocalCredentials)  { $Cred.LocalCredentials }  else { $empty }
                RemoteCredentials = if ($Cred.RemoteCredentials) { $Cred.RemoteCredentials } else { $empty }
            }
        }

        # Single PSCredential - use for both
        if ($Cred -is [System.Management.Automation.PSCredential]) {
            return @{ LocalCredentials = $Cred; RemoteCredentials = $Cred }
        }

        throw "Unsupported Credential type. Pass a PSCredential or the object returned by -SetCredentials."
    }

    function Mount-ShareDrive {
        param([string]$Path, [System.Management.Automation.PSCredential]$Cred)
        $empty = [System.Management.Automation.PSCredential]::Empty
        if ($Cred -eq $empty) { return $null }

        $name = "SpeedTest_$(Get-Random)"
        $drive = New-PSDrive -Name $name -PSProvider FileSystem `
            -Root $Path -Credential $Cred -ErrorAction Stop
        Write-Verbose "Mounted $Path as ${name}:"
        return $drive
    }

    # -------------------------------------------------------------------------
    # SET-CREDENTIALS MODE
    # -------------------------------------------------------------------------

    if ($PSCmdlet.ParameterSetName -eq 'SetCreds') {
        Write-Host ""
        Write-Host "  Network Speed Test - Credential Setup" -ForegroundColor Yellow
        Write-Host "  --------------------------------------" -ForegroundColor DarkGray

        $local = Invoke-CredentialPrompt -Label 'LocalCredentials (Local / Source share)'

        $remote = Invoke-CredentialPrompt -Label 'RemoteCredentials (Remote / Target share)' `
            -CopyKey '2' -CopyLabel 'LocalCredentials'

        if ($null -eq $remote) { $remote = $local }

        Write-Host ""
        Write-Host "  Credentials stored." -ForegroundColor Green

        return [PSCustomObject]@{
            LocalCredentials  = $local
            RemoteCredentials = $remote
        }
    }

    # -------------------------------------------------------------------------
    # HELPERS
    # -------------------------------------------------------------------------

    function New-FailedResult {
        param([string]$Server, [string]$Reason)
        [PSCustomObject]@{
            Server      = $Server
            Timestamp   = Get-Date
            TotalSizeMB = 0
            WriteMbps   = 0
            WriteTime   = [TimeSpan]::Zero
            ReadMbps    = 0
            ReadTime    = [TimeSpan]::Zero
            Status      = "FAILED: $Reason"
        }
    }

    function New-RandomFiles {
        param([string]$Path, [int]$Count, [int]$SizeMB)
        $bytes = $SizeMB * 1MB
        $rng   = [System.Random]::new()
        $buf   = [byte[]]::new([Math]::Min($bytes, 4MB))  # write in 4MB chunks

        for ($i = 1; $i -le $Count; $i++) {
            $filePath = Join-Path $Path "speedtest_$i.tmp"
            $stream   = [System.IO.File]::OpenWrite($filePath)
            $written  = 0
            try {
                while ($written -lt $bytes) {
                    $chunk = [Math]::Min($buf.Length, $bytes - $written)
                    $rng.NextBytes($buf)
                    $stream.Write($buf, 0, $chunk)
                    $written += $chunk
                }
            }
            finally { $stream.Close() }
            Write-Verbose "Generated: $(Split-Path $filePath -Leaf) ($SizeMB MB)"
        }
    }

    # -------------------------------------------------------------------------
    # MAIN RUN
    # -------------------------------------------------------------------------

    # $server extracted early so New-FailedResult is usable during validation
    $server = if ($RemotePath -match '^\\\\([^\\]+)') { $Matches[1] } else { $RemotePath }

    # --- Input validation: all errors returned as structured result objects, no red text ---

    # LocalPath: empty, file-vs-folder, existence, read access
    if ([string]::IsNullOrWhiteSpace($LocalPath)) {
        return New-FailedResult -Server $server -Reason "LocalPath must not be empty."
    }
    if (Test-Path $LocalPath -PathType Leaf) {
        return New-FailedResult -Server $server -Reason "LocalPath '$LocalPath' is a file, not a folder. Provide a directory path."
    }
    if (-not (Test-Path $LocalPath -PathType Container)) {
        if ($GenerateFiles) {
            # Auto-create when generating files - the folder is part of the test setup
            try {
                New-Item -Path $LocalPath -ItemType Directory -ErrorAction Stop | Out-Null
                Write-Verbose "Created LocalPath: $LocalPath"
            } catch {
                return New-FailedResult -Server $server -Reason "LocalPath '$LocalPath' does not exist and could not be created: $_"
            }
        } else {
            return New-FailedResult -Server $server -Reason "LocalPath '$LocalPath' does not exist. Create the folder first, or add -GenerateFiles to create it automatically."
        }
    } else {
        try {
            Get-ChildItem -Path $LocalPath -ErrorAction Stop | Out-Null
        } catch {
            return New-FailedResult -Server $server -Reason "LocalPath '$LocalPath' exists but is not readable (permission denied or locked): $_"
        }
    }

    # RemotePath: UNC format, completeness, host reachability, share accessibility
    if (-not $RemotePath.StartsWith('\\')) {
        return New-FailedResult -Server $server -Reason "RemotePath '$RemotePath' is not a valid UNC path. Expected format: \\server\share"
    }
    $uncParts = $RemotePath.TrimStart('\') -split '\\'
    if ($uncParts.Count -lt 2 -or [string]::IsNullOrWhiteSpace($uncParts[0]) -or [string]::IsNullOrWhiteSpace($uncParts[1])) {
        return New-FailedResult -Server $server -Reason "RemotePath '$RemotePath' is incomplete. Expected format: \\server\share"
    }
    $remoteHost = $uncParts[0]
    if (-not (Test-Connection -ComputerName $remoteHost -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
        return New-FailedResult -Server $remoteHost -Reason "Remote host '$remoteHost' is unreachable (ping failed). Check hostname and network connectivity."
    }
    try {
        $null = Get-Item -Path $RemotePath -ErrorAction Stop
    } catch [System.UnauthorizedAccessException] {
        return New-FailedResult -Server $remoteHost -Reason "Access denied to '$RemotePath'. Provide credentials via -Credential or -SetCredentials."
    } catch {
        return New-FailedResult -Server $remoteHost -Reason "Remote share '$RemotePath' is not accessible: $_"
    }

    $server    = ($RemotePath -split '\\')[2]
    $sessionId = [System.Guid]::NewGuid().ToString('N').Substring(0, 8)  # short unique token
    $targetDir = Join-Path $RemotePath "spd_$sessionId"     # remote write target (A->B)
    $readStage = Join-Path $LocalPath  "spd_${sessionId}_r" # local read target   (B->A)
    $srcDir    = $LocalPath                                  # default source dir
    $timestamp = Get-Date
    $creds     = Resolve-Credentials -Cred $Credential
    $localDrive  = $null
    $remoteDrive = $null

    try {
        # --- Mount shares ---
        $localDrive  = Mount-ShareDrive -Path $LocalPath  -Cred $creds.LocalCredentials
        $remoteDrive = Mount-ShareDrive -Path $RemotePath -Cred $creds.RemoteCredentials

        # --- Generate test files into their own local subdir ---
        if ($GenerateFiles) {
            $srcDir = Join-Path $LocalPath "spd_${sessionId}_w"
            New-Item -Path $srcDir -ItemType Directory -ErrorAction Stop | Out-Null
            Write-Verbose "Generating $FileCount x ${FileSizeMB} MB test files..."
            New-RandomFiles -Path $srcDir -Count $FileCount -SizeMB $FileSizeMB
        }

        # --- Enumerate source files ---
        $files = @(Get-ChildItem -Path $srcDir -File -ErrorAction Stop)
        if ($files.Count -eq 0) {
            return New-FailedResult -Server $server -Reason "No files found in '$srcDir'"
        }

        $totalBytes  = [long]($files | Measure-Object -Property Length -Sum).Sum
        $totalSizeMB = [Math]::Round($totalBytes / 1048576.0, 3)
        Write-Verbose "$($files.Count) file(s) | $totalSizeMB MB total"

        # --- Create remote and local staging dirs ---
        New-Item -Path $targetDir -ItemType Directory -ErrorAction Stop | Out-Null
        Write-Verbose "Created remote dir: $targetDir"
        New-Item -Path $readStage -ItemType Directory -ErrorAction Stop | Out-Null
        Write-Verbose "Created local read stage: $readStage"

        # --- Write test: A->B ---
        Write-Verbose "Write test (A->B)..."
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        foreach ($file in $files) {
            Copy-Item -Path $file.FullName -Destination $targetDir -Force -ErrorAction Stop
        }
        $sw.Stop()
        $writeSeconds = $sw.Elapsed.TotalSeconds

        # --- Read test: B->A (into clean local staging dir, not source) ---
        Write-Verbose "Read test (B->A)..."
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        foreach ($file in $files) {
            Copy-Item -Path (Join-Path $targetDir $file.Name) `
                      -Destination $readStage -Force -ErrorAction Stop
        }
        $sw.Stop()
        $readSeconds = $sw.Elapsed.TotalSeconds

        # Proven formula: (bytes * 8 bits) / seconds / 1048576 bits-per-Mbit
        $writeMbps = [Math]::Round((([double]$totalBytes * 8) / $writeSeconds) / 1048576, 2)
        $readMbps  = [Math]::Round((([double]$totalBytes * 8) / $readSeconds)  / 1048576, 2)

        Write-Verbose "Write: $writeMbps Mbps | Read: $readMbps Mbps"

        return [PSCustomObject]@{
            Server      = $server
            Timestamp   = $timestamp
            TotalSizeMB = $totalSizeMB
            WriteMbps   = $writeMbps
            WriteTime   = [TimeSpan]::FromSeconds($writeSeconds)
            ReadMbps    = $readMbps
            ReadTime    = [TimeSpan]::FromSeconds($readSeconds)
            Status      = 'OK'
        }
    }
    catch {
        return New-FailedResult -Server $server -Reason $_
    }
    finally {
        # Cleanup all session-scoped dirs (remote write target, local read stage, local source if generated)
        foreach ($dir in @($targetDir, $readStage, $(if ($GenerateFiles) { $srcDir })) | Where-Object { $_ -and (Test-Path $_) }) {
            Remove-Item -Path $dir -Recurse -Force -ErrorAction SilentlyContinue
            Write-Verbose "Removed: $dir"
        }
        # Unmount PSDrives
        foreach ($drive in @($localDrive, $remoteDrive) | Where-Object { $_ }) {
            Remove-PSDrive -Name $drive.Name -Force -ErrorAction SilentlyContinue
            Write-Verbose "Unmounted drive $($drive.Name)"
        }
    }
}

<#
# 1 — No credentials, existing files
Test-NetworkSpeed -LocalPath ".\files" -RemotePath "\\server1\share"

# 2 — Generate 10×10MB files on the fly (default), no credentials
Test-NetworkSpeed -LocalPath "C:\Temp\test" -RemotePath "\\server1\share" -GenerateFiles

# 3 — Generate 5×50MB files
Test-NetworkSpeed -LocalPath "C:\Temp\test" -RemotePath "\\server1\share" -GenerateFiles -FileCount 5 -FileSizeMB 50

# 4 — Interactive credential setup, then run
$creds = Test-NetworkSpeed -SetCredentials
Test-NetworkSpeed -LocalPath "C:\Temp\test" -RemotePath "\\server1\share" -Credential $creds -GenerateFiles

# 5 — Single flat credential (same for both sides)
$cred = Get-Credential
Test-NetworkSpeed -LocalPath "C:\Temp\test" -RemotePath "\\server1\share" -Credential $cred -GenerateFiles
#>
