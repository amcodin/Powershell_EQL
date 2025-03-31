# Main.ps1

#requires -Version 3.0

# --- Load Assemblies FIRST ---
Write-Host "Attempting to load .NET Assemblies..."
try {
    Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Write-Host ".NET Assemblies loaded successfully."
} catch {
    Write-Error "FATAL: Failed to load required .NET Assemblies. UI cannot be created."
    Write-Error "Error Message: $($_.Exception.Message)"
    if ($Host.Name -eq 'ConsoleHost' -and -not $psISE -and $env:TERM_PROGRAM -ne 'vscode') { Read-Host "Press Enter to exit" }
    Exit 1
}

#region Global Variables & Constants
$script:DefaultPath = "C:\LocalData"
#endregion

#region Functions

# --- Core Logic Functions (Refactored) ---

function Invoke-BackupOperation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$BackupRootPath,

        [Parameter(Mandatory=$true)]
        [System.Collections.Generic.List[PSCustomObject]]$ItemsToBackup,

        [Parameter(Mandatory=$true)]
        [bool]$BackupNetworkDrives,

        [Parameter(Mandatory=$true)]
        [bool]$BackupPrinters,

        [Parameter(Mandatory=$false)]
        [scriptblock]$ProgressAction # Optional scriptblock for progress updates: $ProgressAction.Invoke($status, $percent, $details)
    )

    Write-Host "--- Starting Backup Operation to '$BackupRootPath' ---"
    $UpdateProgress = { if ($ProgressAction) { $ProgressAction.Invoke($args[0], $args[1], $args[2]) } else { Write-Host "$($args[0]) ($($args[1])%) - $($args[2])" } }

    try { # Line 48 - Start of try block
        $UpdateProgress.Invoke("Initializing Backup", 0, "Creating backup directory...")
        if (-not (Test-Path $BackupRootPath)) {
            New-Item -Path $BackupRootPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Write-Host "Backup directory created: $BackupRootPath"
        }

        $csvLogPath = Join-Path $BackupRootPath "FileList_Backup.csv"
        Write-Host "Creating log file: $csvLogPath"
        "OriginalFullPath,BackupRelativePath" | Set-Content -Path $csvLogPath -Encoding UTF8

        if (-not $ItemsToBackup -or $ItemsToBackup.Count -eq 0) {
             throw "No items specified for backup."
        }
        Write-Host "Processing $($ItemsToBackup.Count) items for backup."

        # Estimate total files for progress (more accurate calculation)
        $UpdateProgress.Invoke("Estimating size...", 0, "Calculating total files...")
        $totalFilesEstimate = 0
        $itemsToProcessForSize = $ItemsToBackup # Use the provided list
        $itemsToProcessForSize | ForEach-Object {
            if (Test-Path -LiteralPath $_.Path -ErrorAction SilentlyContinue) {
                if ($_.Type -eq 'Folder') {
                    try { $totalFilesEstimate += (Get-ChildItem -LiteralPath $_.Path -Recurse -File -Force -ErrorAction SilentlyContinue).Count } catch {}
                } else { $totalFilesEstimate++ }
            }
        }
        # --- FIX: Replace ternary operator ---
        $networkDriveCount = $(if ($BackupNetworkDrives) { 1 } else { 0 })
        $printerCount = $(if ($BackupPrinters) { 1 } else { 0 })
        if ($BackupNetworkDrives) { $totalFilesEstimate++ } # Keep this for file count estimate
        if ($BackupPrinters) { $totalFilesEstimate++ } # Keep this for file count estimate
        Write-Host "Estimated total files/items: $totalFilesEstimate"

        $currentItemIndex = 0
        # --- FIX: Calculate totalItems using compatible logic ---
        $totalItems = $ItemsToBackup.Count + $networkDriveCount + $printerCount

        # Process Files/Folders
        foreach ($item in $ItemsToBackup) {
            $currentItemIndex++
            $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 100) } else { 0 }
            $statusMessage = "Backing up Item $currentItemIndex of $totalItems"
            $UpdateProgress.Invoke($statusMessage, $percentComplete, "Processing: $($item.Name)")
            Write-Host "Processing item: $($item.Name) ($($item.Type)) - Path: $($item.Path)"
            $sourcePath = $item.Path

            if (-not (Test-Path -LiteralPath $sourcePath)) {
                Write-Warning "Source path not found, skipping: $sourcePath"
                continue
            }

            if ($item.Type -eq "Folder") {
                Write-Host "  Item is a folder. Processing recursively..."
                try {
                    $basePathLength = $sourcePath.TrimEnd('\').Length
                    Get-ChildItem -LiteralPath $sourcePath -Recurse -File -Force -ErrorAction Stop | ForEach-Object {
                        $originalFileFullPath = $_.FullName
                        $relativeFilePath = $originalFileFullPath.Substring($basePathLength).TrimStart('\')
                        $backupRelativePath = Join-Path $item.Name $relativeFilePath
                        $targetBackupPath = Join-Path $BackupRootPath $backupRelativePath
                        $targetBackupDir = [System.IO.Path]::GetDirectoryName($targetBackupPath)

                        if (-not (Test-Path $targetBackupDir)) {
                            New-Item -Path $targetBackupDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
                        }
                        Copy-Item -LiteralPath $originalFileFullPath -Destination $targetBackupPath -Force -ErrorAction Stop
                        "`"$originalFileFullPath`",`"$backupRelativePath`"" | Add-Content -Path $csvLogPath -Encoding UTF8
                        $UpdateProgress.Invoke($statusMessage, $percentComplete, "Copied: $($_.Name)")
                    }
                    Write-Host "  Finished processing folder: $($item.Name)"
                } catch { Write-Warning "Error processing folder '$($item.Name)' ($sourcePath): $($_.Exception.Message)" }
            } else { # Single File
                 Write-Host "  Item is a file. Processing..."
                 try {
                    $originalFileFullPath = $sourcePath
                    $backupRelativePath = $item.Name
                    $targetBackupPath = Join-Path $BackupRootPath $backupRelativePath
                    Copy-Item -LiteralPath $originalFileFullPath -Destination $targetBackupPath -Force -ErrorAction Stop
                    "`"$originalFileFullPath`",`"$backupRelativePath`"" | Add-Content -Path $csvLogPath -Encoding UTF8
                 } catch { Write-Warning "Error processing file '$($item.Name)' ($sourcePath): $($_.Exception.Message)" }
            }
        } # End foreach item

        # Backup Network Drives
        if ($BackupNetworkDrives) {
            $currentItemIndex++
            $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 100) } else { 0 }
            $statusMessage = "Backing up Item $currentItemIndex of $totalItems"
            $UpdateProgress.Invoke($statusMessage, $percentComplete, "Backing up network drives...")
            Write-Host "Processing Network Drives backup..."
            try {
                Get-WmiObject -Class Win32_MappedLogicalDisk -ErrorAction Stop |
                    Select-Object Name, ProviderName |
                    Export-Csv -Path (Join-Path $BackupRootPath "Drives.csv") -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
                Write-Host "Network drives backed up successfully." -ForegroundColor Green
            } catch { Write-Warning "Failed to backup network drives: $($_.Exception.Message)" }
        }

        # Backup Printers
        if ($BackupPrinters) {
            $currentItemIndex++
            $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 100) } else { 0 }
            $statusMessage = "Backing up Item $currentItemIndex of $totalItems"
            $UpdateProgress.Invoke($statusMessage, $percentComplete, "Backing up printers...")
            Write-Host "Processing Printers backup..."
            try {
                Get-WmiObject -Class Win32_Printer -Filter "Local = False" -ErrorAction Stop |
                    Select-Object -ExpandProperty Name |
                    Set-Content -Path (Join-Path $BackupRootPath "Printers.txt") -Encoding UTF8 -ErrorAction Stop
                 Write-Host "Printers backed up successfully." -ForegroundColor Green
            } catch { Write-Warning "Failed to backup printers: $($_.Exception.Message)" }
        }

        $UpdateProgress.Invoke("Backup Complete", 100, "Successfully backed up to: $BackupRootPath")
        Write-Host "--- Backup Operation Finished ---"
        return $true # Indicate success

    } catch { # Line 163 - Catch for the try block started on Line 48
        $errorMessage = "Backup Operation Failed: $($_.Exception.Message)"
        Write-Error $errorMessage
        $UpdateProgress.Invoke("Backup Failed", -1, $errorMessage) # Use -1 percent for error state
        return $false # Indicate failure
    }
} # Line 169 - End of Invoke-BackupOperation function

function Invoke-RestoreOperation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$BackupRootPath,

        [Parameter(ParameterSetName='AllFromLog', Mandatory=$true)]
        [switch]$RestoreAllFromLog,

        [Parameter(Mandatory=$true)]
        [bool]$RestoreNetworkDrives,

        [Parameter(Mandatory=$true)]
        [bool]$RestorePrinters,

        [Parameter(Mandatory=$false)]
        [scriptblock]$ProgressAction
    )

    Write-Host "--- Starting Restore Operation from '$BackupRootPath' ---"
    $UpdateProgress = { if ($ProgressAction) { $ProgressAction.Invoke($args[0], $args[1], $args[2]) } else { Write-Host "$($args[0]) ($($args[1])%) - $($args[2])" } }

    try {
        $UpdateProgress.Invoke("Initializing Restore", 0, "Checking backup contents...")
        $csvLogPath = Join-Path $BackupRootPath "FileList_Backup.csv"
        $drivesCsvPath = Join-Path $BackupRootPath "Drives.csv"
        $printersTxtPath = Join-Path $BackupRootPath "Printers.txt"

        if (-not (Test-Path $csvLogPath -PathType Leaf)) {
            throw "Backup log file 'FileList_Backup.csv' not found in: $BackupRootPath"
        }

        Write-Host "Importing backup log file..."
        $backupLog = Import-Csv -Path $csvLogPath -Encoding UTF8
        if (-not $backupLog) {
             throw "Backup log file is empty or could not be read: $csvLogPath"
        }
        Write-Host "Imported $($backupLog.Count) entries from log file."

        $logEntriesToRestore = $null
        if ($RestoreAllFromLog) {
            $logEntriesToRestore = $backupLog
            Write-Host "RestoreAllFromLog specified: Processing all $($logEntriesToRestore.Count) log entries."
        } else {
             throw "Selective restore not implemented in this refactoring yet."
        }

        if (-not $logEntriesToRestore) {
            throw "No log entries identified for restore."
        }

        # Estimate progress
        # --- FIX: Replace ternary operator ---
        $networkDriveCount = $(if ($RestoreNetworkDrives) { 1 } else { 0 })
        $printerCount = $(if ($RestorePrinters) { 1 } else { 0 })
        $totalItems = $logEntriesToRestore.Count + $networkDriveCount + $printerCount
        $currentItemIndex = 0
        Write-Host "Total items to restore (files/folders + drives + printers): $totalItems"

        # Restore Files/Folders from Log
        Write-Host "Starting restore of files/folders..."
        foreach ($entry in $logEntriesToRestore) {
            $currentItemIndex++
            $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 100) } else { 0 }
            $statusMessage = "Restoring Item $currentItemIndex of $totalItems"

            $originalFileFullPath = $entry.OriginalFullPath
            $backupRelativePath = $entry.BackupRelativePath
            $sourceBackupPath = Join-Path $BackupRootPath $backupRelativePath

            $UpdateProgress.Invoke($statusMessage, $percentComplete, "Restoring: $(Split-Path $originalFileFullPath -Leaf)")

            if (Test-Path -LiteralPath $sourceBackupPath -PathType Leaf) {
                try {
                    $targetRestoreDir = [System.IO.Path]::GetDirectoryName($originalFileFullPath)
                    if (-not (Test-Path $targetRestoreDir)) {
                        New-Item -Path $targetRestoreDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
                    }
                    Copy-Item -LiteralPath $sourceBackupPath -Destination $originalFileFullPath -Force -ErrorAction Stop
                } catch { Write-Warning "Failed to restore '$originalFileFullPath' from '$sourceBackupPath': $($_.Exception.Message)" }
            } else { Write-Warning "Source file not found in backup, skipping restore: $sourceBackupPath (Expected for: $originalFileFullPath)" }
        }
        Write-Host "Finished restoring files/folders from log."

        # Restore Network Drives
        if ($RestoreNetworkDrives) {
            $currentItemIndex++
            $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 100) } else { 0 }
            $statusMessage = "Restoring Item $currentItemIndex of $totalItems"
            $UpdateProgress.Invoke($statusMessage, $percentComplete, "Restoring network drives...")
            Write-Host "Processing Network Drives restore..."
            if (Test-Path $drivesCsvPath) {
                Write-Host "Found Drives.csv. Processing mappings..."
                try {
                    Import-Csv $drivesCsvPath | ForEach-Object {
                        $driveLetter = $_.Name.TrimEnd(':')
                        $networkPath = $_.ProviderName
                        if ($driveLetter -match '^[A-Z]$' -and $networkPath -match '^\\\\') {
                            if (-not (Test-Path -LiteralPath "$($driveLetter):")) {
                                try {
                                    Write-Host "  Mapping $driveLetter to $networkPath"
                                    New-PSDrive -Name $driveLetter -PSProvider FileSystem -Root $networkPath -Persist -Scope Global -ErrorAction Stop
                                } catch { Write-Warning "  Failed to map drive $driveLetter`: $($_.Exception.Message)" }
                            } else { Write-Host "  Drive $driveLetter already exists, skipping." }
                        } else { Write-Warning "  Skipping invalid drive mapping: Name='$($_.Name)', Provider='$networkPath'" }
                    }
                    Write-Host "Finished processing network drive mappings."
                } catch { Write-Warning "Error processing network drive restorations: $($_.Exception.Message)" }
            } else { Write-Warning "Network drives backup file (Drives.csv) not found." }
        }

        # Restore Printers
        if ($RestorePrinters) {
            $currentItemIndex++
            $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 100) } else { 0 }
            $statusMessage = "Restoring Item $currentItemIndex of $totalItems"
            $UpdateProgress.Invoke($statusMessage, $percentComplete, "Restoring printers...")
            Write-Host "Processing Printers restore..."
            if (Test-Path $printersTxtPath) {
                 Write-Host "Found Printers.txt. Processing printers..."
                try {
                    $wsNet = New-Object -ComObject WScript.Network
                    Get-Content $printersTxtPath | ForEach-Object {
                        $printerPath = $_.Trim()
                        if (-not ([string]::IsNullOrWhiteSpace($printerPath)) -and $printerPath -match '^\\\\') {
                            Write-Host "  Attempting to add printer: $printerPath"
                            try {
                                 $wsNet.AddWindowsPrinterConnection($printerPath)
                                 Write-Host "    Added printer connection (or it already existed)."
                            } catch { Write-Warning "    Failed to add printer '$printerPath': $($_.Exception.Message)" }
                        } else { Write-Warning "  Skipping invalid or empty line in Printers.txt: '$_'" }
                    }
                    Write-Host "Finished processing printers."
                } catch { Write-Warning "Error processing printer restorations: $($_.Exception.Message)" }
            } else { Write-Warning "Printers backup file (Printers.txt) not found." }
        }

        $UpdateProgress.Invoke("Restore Complete", 100, "Successfully restored from: $BackupRootPath")
        Write-Host "--- Restore Operation Finished ---"
        return $true # Indicate success

    } catch {
        $errorMessage = "Restore Operation Failed: $($_.Exception.Message)"
        Write-Error $errorMessage
        $UpdateProgress.Invoke("Restore Failed", -1, $errorMessage) # Use -1 percent for error state
        return $false # Indicate failure
    }
}

# --- Background Job Starter ---
function Start-BackgroundUpdateJob {
    Write-Host "Initiating background system updates job..." -ForegroundColor Yellow
    $script:updateJob = Start-Job -Name "BackgroundUpdates" -ScriptBlock {
        # Functions need to be defined within the job's scope
        function Set-GPupdate {
            Write-Host "JOB: Initiating Group Policy update..." -ForegroundColor Cyan
            try {
                $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c gpupdate /force" -PassThru -Wait -ErrorAction Stop
                if ($process.ExitCode -eq 0) { Write-Host "JOB: Group Policy update completed successfully." -ForegroundColor Green }
                else { Write-Warning "JOB: Group Policy update process finished with exit code: $($process.ExitCode)." }
            } catch { Write-Error "JOB: Failed to start GPUpdate process: $($_.Exception.Message)" }
            Write-Host "JOB: Exiting Set-GPupdate function."
        }
        function Start-ConfigManagerActions {
             param()
             Write-Host "JOB: Entering Start-ConfigManagerActions function."
             $ccmExecPath = "C:\Windows\CCM\ccmexec.exe"; $clientSDKNamespace = "root\ccm\clientsdk"; $clientClassName = "CCM_ClientUtilities"; $scheduleMethodName = "TriggerSchedule"; $overallSuccess = $false; $cimAttemptedAndSucceeded = $false
             $scheduleActions = @( @{ID = '{00000000-0000-0000-0000-000000000021}'; Name = 'Machine Policy Retrieval & Evaluation Cycle'}, @{ID = '{00000000-0000-0000-0000-000000000022}'; Name = 'User Policy Retrieval & Evaluation Cycle'}, @{ID = '{00000000-0000-0000-0000-000000000001}'; Name = 'Hardware Inventory Cycle'}, @{ID = '{00000000-0000-0000-0000-000000000002}'; Name = 'Software Inventory Cycle'}, @{ID = '{00000000-0000-0000-0000-000000000113}'; Name = 'Software Updates Scan Cycle'}, @{ID = '{00000000-0000-0000-0000-000000000101}'; Name = 'Hardware Inventory Collection Cycle'}, @{ID = '{00000000-0000-0000-0000-000000000108}'; Name = 'Software Updates Assignments Evaluation Cycle'}, @{ID = '{00000000-0000-0000-0000-000000000102}'; Name = 'Software Inventory Collection Cycle'} )
             Write-Host "JOB: Defined $($scheduleActions.Count) CM actions to trigger."
             $ccmService = Get-Service -Name CcmExec -ErrorAction SilentlyContinue
             if (-not $ccmService) { Write-Warning "JOB: CM service (CcmExec) not found. Skipping."; return $false }
             elseif ($ccmService.Status -ne 'Running') { Write-Warning "JOB: CM service (CcmExec) is not running (Status: $($ccmService.Status)). Skipping."; return $false }
             else { Write-Host "JOB: CM service (CcmExec) found and running." }
             Write-Host "JOB: Attempting Method 1: Triggering via CIM..."
             $cimMethodSuccess = $true
             try {
                 if (Get-CimClass -Namespace $clientSDKNamespace -ClassName $clientClassName -ErrorAction SilentlyContinue) {
                      Write-Host "JOB: CIM Class found."
                      foreach ($action in $scheduleActions) {
                         Write-Host "JOB:   Triggering $($action.Name) (ID: $($action.ID)) via CIM."
                         try { Invoke-CimMethod -Namespace $clientSDKNamespace -ClassName $clientClassName -MethodName $scheduleMethodName -Arguments @{sScheduleID = $action.ID} -ErrorAction Stop; Write-Host "JOB:     $($action.Name) triggered successfully via CIM." }
                         catch { Write-Warning "JOB:     Failed to trigger $($action.Name) via CIM: $($_.Exception.Message)"; $cimMethodSuccess = $false }
                      }
                      if ($cimMethodSuccess) { $cimAttemptedAndSucceeded = $true; $overallSuccess = $true; Write-Host "JOB: All actions successfully triggered via CIM." -ForegroundColor Green }
                      else { Write-Warning "JOB: One or more actions failed to trigger via CIM." }
                 } else { Write-Warning "JOB: CIM Class '$clientClassName' not found. Cannot use CIM method."; $cimMethodSuccess = $false }
             } catch { Write-Error "JOB: An unexpected error occurred during CIM attempt: $($_.Exception.Message)"; $cimMethodSuccess = $false }
             if (-not $cimAttemptedAndSucceeded) {
                 Write-Host "JOB: CIM failed/unavailable. Attempting Method 2: Fallback via ccmexec.exe..."
                 if (Test-Path -Path $ccmExecPath -PathType Leaf) {
                     Write-Host "JOB: Found $ccmExecPath."
                     $execMethodSuccess = $true
                     foreach ($action in $scheduleActions) {
                         Write-Host "JOB:   Triggering $($action.Name) (ID: $($action.ID)) via ccmexec.exe."
                         try { $process = Start-Process -FilePath $ccmExecPath -ArgumentList "-TriggerSchedule $($action.ID)" -NoNewWindow -PassThru -Wait -ErrorAction Stop; if ($process.ExitCode -ne 0) { Write-Warning "JOB:     $($action.Name) via ccmexec.exe finished with exit code $($process.ExitCode)." } else { Write-Host "JOB:     $($action.Name) triggered via ccmexec.exe (Exit Code 0)." } }
                         catch { Write-Warning "JOB:     Failed to execute ccmexec.exe for $($action.Name): $($_.Exception.Message)"; $execMethodSuccess = $false }
                     }
                     if ($execMethodSuccess) { $overallSuccess = $true; Write-Host "JOB: Finished attempting actions via ccmexec.exe." -ForegroundColor Green }
                     else { Write-Warning "JOB: One or more actions failed to execute via ccmexec.exe." }
                 } else { Write-Warning "JOB: Fallback executable not found at $ccmExecPath." }
             }
             if ($overallSuccess) { Write-Host "JOB: CM actions attempt finished successfully." -ForegroundColor Green }
             else { Write-Warning "JOB: CM actions attempt finished, but could not be confirmed as fully successful." }
             Write-Host "JOB: Exiting Start-ConfigManagerActions function."
             return $overallSuccess
        }
        # Execute functions
        Set-GPupdate
        Start-ConfigManagerActions
        Write-Host "JOB: Background updates finished."
    }
    Write-Host "Background update job started (ID: $($script:updateJob.Id)). Output will be shown later." -ForegroundColor Yellow
}


# --- Data Gathering Functions ---
function Get-BackupPaths {
    [CmdletBinding()]
    param ()
    $specificPaths = @( "$env:APPDATA\Microsoft\Signatures", 
    "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms", 
    "$env:APPDATA\Microsoft\Sticky Notes\StickyNotes.snt", 
    "$env:LOCALAPPDATA\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\plum.sqlite", 
    "$env:APPDATA\google\googleearth\myplaces.kml", 
    "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Bookmarks",
    "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Bookmarks" )
    $result = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($path in $specificPaths) {
        $resolvedPath = try { $ExecutionContext.InvokeCommand.ExpandString($path) } catch { $null }
        if ($resolvedPath -and (Test-Path -LiteralPath $resolvedPath)) {
            $item = Get-Item -LiteralPath $resolvedPath -ErrorAction SilentlyContinue
            if ($item) { $result.Add([PSCustomObject]@{ Name = $item.Name; Path = $item.FullName; Type = if ($item.PSIsContainer) { "Folder" } else { "File" } }) }
            else { Write-Host "Specific path resolved but Get-Item failed: $resolvedPath" }
        } else { Write-Host "Specific path not found or resolution failed: $path" }
    }
    return $result
}

function Get-UserPaths {
    [CmdletBinding()]
    param ()
    $specificPaths = @( 
        "$env:USERPROFILE\Downloads",
        "$env:USERPROFILE\Pictures",
        "$env:USERPROFILE\Videos")
    $result = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($path in $specificPaths) {
        $resolvedPath = try { $ExecutionContext.InvokeCommand.ExpandString($path) } catch { $null }
        if ($resolvedPath -and (Test-Path -LiteralPath $resolvedPath -PathType Container)) {
            $item = Get-Item -LiteralPath $resolvedPath -ErrorAction SilentlyContinue
             if ($item) { $result.Add([PSCustomObject]@{ Name = $item.Name; Path = $item.FullName; Type = "Folder" }) }
             else { Write-Host "BAU path resolved but Get-Item failed: $resolvedPath" }
        } else { Write-Host "BAU path not found, not a folder, or resolution failed: $path" }
    }
    return $result
}

# --- Helper Function ---
function Format-Bytes {
    param([Parameter(Mandatory=$true)][long]$Bytes)
    $suffix = @("B", "KB", "MB", "GB", "TB", "PB"); $index = 0; $value = [double]$Bytes
    while ($value -ge 1024 -and $index -lt ($suffix.Length - 1)) { $value /= 1024; $index++ }
    return "{0:N2} {1}" -f $value, $suffix[$index]
}

# --- GUI Functions ---

# Show mode selection dialog
function Show-ModeDialog {
    Write-Host "Entering Show-ModeDialog function."
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Operation"
    $form.Size = New-Object System.Drawing.Size(400, 150) # Wider for 3 buttons
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false; $form.MinimizeBox = $false; $form.HelpButton = $false

    $btnBackup = New-Object System.Windows.Forms.Button
    $btnBackup.Location = New-Object System.Drawing.Point(30, 40)
    $btnBackup.Size = New-Object System.Drawing.Size(80, 30)
    $btnBackup.Text = "Backup"
    $btnBackup.DialogResult = [System.Windows.Forms.DialogResult]::Yes # Using Yes for Backup
    $form.Controls.Add($btnBackup)

    $btnRestore = New-Object System.Windows.Forms.Button
    $btnRestore.Location = New-Object System.Drawing.Point(140, 40)
    $btnRestore.Size = New-Object System.Drawing.Size(80, 30)
    $btnRestore.Text = "Restore"
    $btnRestore.DialogResult = [System.Windows.Forms.DialogResult]::No # Using No for Restore
    $form.Controls.Add($btnRestore)

    # NEW Express Button
    $btnExpress = New-Object System.Windows.Forms.Button
    $btnExpress.Location = New-Object System.Drawing.Point(250, 40)
    $btnExpress.Size = New-Object System.Drawing.Size(80, 30)
    $btnExpress.Text = "Express"
    $btnExpress.DialogResult = [System.Windows.Forms.DialogResult]::OK # Using OK for Express
    $form.Controls.Add($btnExpress)

    # Set default button (optional, maybe Express?)
    $form.AcceptButton = $btnExpress
    # Allow closing via Esc key (maps to Cancel)
    $form.CancelButton = $btnRestore # Or add a dedicated Cancel button

    Write-Host "Showing mode selection dialog."
    $result = $form.ShowDialog()
    $form.Dispose()
    Write-Host "Mode selection dialog closed with result: $result"

    # Determine mode based on DialogResult
    $selectedMode = switch ($result) {
        ([System.Windows.Forms.DialogResult]::Yes) { 'Backup' }
        ([System.Windows.Forms.DialogResult]::No) { 'Restore' }
        ([System.Windows.Forms.DialogResult]::OK) { 'Express' }
        Default { 'Cancel' } # Or $null
    }

    Write-Host "Determined mode: $selectedMode" -ForegroundColor Cyan
    return $selectedMode
}

# Show main window (Modified to use Invoke- functions)
# Show main window (Modified to use Invoke- functions)
function Show-MainWindow {
    param(
        [Parameter(Mandatory=$true)]
        [bool]$IsBackup
    )
    $modeString = if ($IsBackup) { 'Backup' } else { 'Restore' }
    Write-Host "Entering Show-MainWindow function. Mode: $modeString"

    # XAML UI Definition (remains the same as previous version)
    [xml]$XAML = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="User Data Backup/Restore Tool"
    Width="800"
    Height="650"
    WindowStartupLocation="CenterScreen">
    <Grid>
        <Label Content="Location:" Margin="10,10,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <TextBox Name="txtSaveLoc" Width="400" Height="30" Margin="10,40,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" IsReadOnly="True"/>
        <Button Name="btnBrowse" Content="Browse" Width="60" Height="30" Margin="420,40,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <Label Name="lblMode" Content="" Margin="500,10,10,0" HorizontalAlignment="Right" VerticalAlignment="Top" FontWeight="Bold"/>
        <Label Name="lblFreeSpace" Content="Free Space: -" Margin="500,40,10,0" HorizontalAlignment="Right" VerticalAlignment="Top"/>
        <Label Name="lblRequiredSpace" Content="Required Space: -" Margin="500,65,10,0" HorizontalAlignment="Right" VerticalAlignment="Top"/>
        <Label Name="lblStatus" Content="Ready" Margin="10,0,10,10" HorizontalAlignment="Center" VerticalAlignment="Bottom" FontStyle="Italic"/>
        <Label Content="Files/Folders to Process:" Margin="10,90,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <ListView Name="lvwFiles" Margin="10,120,200,140" SelectionMode="Extended">
             <ListView.View>
                <GridView>
                    <GridViewColumn Width="30">
                        <GridViewColumn.CellTemplate>
                            <DataTemplate> <CheckBox IsChecked="{Binding IsSelected, Mode=TwoWay}" /> </DataTemplate>
                        </GridViewColumn.CellTemplate>
                    </GridViewColumn>
                    <GridViewColumn Header="Name" DisplayMemberBinding="{Binding Name}" Width="180"/>
                    <GridViewColumn Header="Type" DisplayMemberBinding="{Binding Type}" Width="70"/>
                    <GridViewColumn Header="Path" DisplayMemberBinding="{Binding Path}" Width="280"/>
                </GridView>
            </ListView.View>
        </ListView>
        <StackPanel Margin="0,120,10,0" HorizontalAlignment="Right" Width="180">
            <Button Name="btnAddFile" Content="Add File" Width="120" Height="30" Margin="0,0,0,10"/>
            <Button Name="btnAddFolder" Content="Add Folder" Width="120" Height="30" Margin="0,0,0,10"/>
            <Button Name="btnAddBAUPaths" Content="Add User Folders" Width="120" Height="30" Margin="0,0,0,10"/>
            <Button Name="btnRemove" Content="Remove Selected" Width="120" Height="30" Margin="0,0,0,20"/>
            <CheckBox Name="chkNetwork" Content="Network Drives" IsChecked="True" Margin="5,0,0,5"/>
            <CheckBox Name="chkPrinters" Content="Printers" IsChecked="True" Margin="5,0,0,5"/>
        </StackPanel>
        <ProgressBar Name="prgProgress" Height="20" Margin="10,0,10,60" VerticalAlignment="Bottom"/>
        <TextBlock Name="txtProgress" Text="" Margin="10,0,10,85" VerticalAlignment="Bottom" TextWrapping="Wrap"/>
        <Button Name="btnStart" Content="Start" Width="100" Height="30" Margin="10,0,0,20" VerticalAlignment="Bottom" HorizontalAlignment="Left" IsDefault="True"/>
    </Grid>
</Window>
'@

    try { # OUTER try for window loading
        Write-Host "Parsing XAML for main window."
        $reader = New-Object System.Xml.XmlNodeReader $XAML
        $window = [Windows.Markup.XamlReader]::Load($reader)
        Write-Host "XAML loaded successfully."
        $window.DataContext = [PSCustomObject]@{ IsRestoreMode = (-not $IsBackup) }

        Write-Host "Finding controls in main window."
        $controls = @{}
        $window.FindName('txtSaveLoc') | ForEach-Object { $controls['txtSaveLoc'] = $_ }
        $window.FindName('btnBrowse') | ForEach-Object { $controls['btnBrowse'] = $_ }
        $window.FindName('btnStart') | ForEach-Object { $controls['btnStart'] = $_ }
        $window.FindName('lblMode') | ForEach-Object { $controls['lblMode'] = $_ }
        $window.FindName('lblStatus') | ForEach-Object { $controls['lblStatus'] = $_ }
        $window.FindName('lvwFiles') | ForEach-Object { $controls['lvwFiles'] = $_ }
        $window.FindName('btnAddFile') | ForEach-Object { $controls['btnAddFile'] = $_ }
        $window.FindName('btnAddFolder') | ForEach-Object { $controls['btnAddFolder'] = $_ }
        $window.FindName('btnRemove') | ForEach-Object { $controls['btnRemove'] = $_ }
        $window.FindName('chkNetwork') | ForEach-Object { $controls['chkNetwork'] = $_ }
        $window.FindName('chkPrinters') | ForEach-Object { $controls['chkPrinters'] = $_ }
        $window.FindName('prgProgress') | ForEach-Object { $controls['prgProgress'] = $_ }
        $window.FindName('txtProgress') | ForEach-Object { $controls['txtProgress'] = $_ }
        $window.FindName('btnAddBAUPaths') | ForEach-Object { $controls['btnAddBAUPaths'] = $_ }
        $window.FindName('lblFreeSpace') | ForEach-Object { $controls['lblFreeSpace'] = $_ }
        $window.FindName('lblRequiredSpace') | ForEach-Object { $controls['lblRequiredSpace'] = $_ }
        Write-Host "Controls found and stored in hashtable."

        # --- Helper Functions for UI Updates (Space Labels) ---
        # Define $script:UpdateFreeSpaceLabel
        $script:UpdateFreeSpaceLabel = {
            param($ControlsParam)
            $location = $ControlsParam.txtSaveLoc.Text
            $freeSpaceString = "Free Space: N/A"
            if (-not [string]::IsNullOrEmpty($location)) {
                try {
                    $driveLetter = $null
                    if ($location -match '^[a-zA-Z]:\\') { # Local path C:\...
                        $driveLetter = $location.Substring(0, 2)
                    } elseif ($location -match '^\\\\[^\\]+\\[^\\]+') { # UNC Path \\server\share\...
                         $freeSpaceString = "Free Space: N/A (UNC)"
                    }

                    if ($driveLetter) {
                        $driveInfo = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$driveLetter'" -ErrorAction SilentlyContinue
                        if ($driveInfo -and $driveInfo.FreeSpace -ne $null) {
                            $freeSpaceString = "Free Space: $(Format-Bytes $driveInfo.FreeSpace)"
                        } else { Write-Warning "Could not get free space for drive $driveLetter" }
                    }
                } catch { Write-Warning "Error getting free space for '$location': $($_.Exception.Message)" }
            }
            $ControlsParam.lblFreeSpace.Content = $freeSpaceString
            Write-Host "Updated Free Space Label: $freeSpaceString"
        }
        # Define $script:UpdateRequiredSpaceLabel
        $script:UpdateRequiredSpaceLabel = {
            param($ControlsParam)
            $totalSize = 0L
            $requiredSpaceString = "Required Space: Calculating..."
            $ControlsParam.lblRequiredSpace.Content = $requiredSpaceString

            $items = @($ControlsParam.lvwFiles.ItemsSource)
            if ($items -ne $null -and $items.Count -gt 0) {
                $checkedItems = $items | Where-Object { $_.IsSelected }
                Write-Host "Calculating required space for $($checkedItems.Count) checked items..."
                foreach ($item in $checkedItems) {
                    if (Test-Path -LiteralPath $item.Path) {
                        try {
                            if ($item.Type -eq 'Folder') {
                                $folderSize = (Get-ChildItem -LiteralPath $item.Path -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                                if ($folderSize -ne $null) { $totalSize += $folderSize }
                            } else { # File
                                $fileSize = (Get-Item -LiteralPath $item.Path -Force -ErrorAction SilentlyContinue).Length
                                if ($fileSize -ne $null) { $totalSize += $fileSize }
                            }
                        } catch { Write-Warning "Error calculating size for '$($item.Path)': $($_.Exception.Message)" }
                    } else { Write-Warning "Checked item path not found, skipping size calculation: $($item.Path)" }
                }
                $requiredSpaceString = "Required Space: $(Format-Bytes $totalSize)"
            } else { $requiredSpaceString = "Required Space: 0 B" }

            $ControlsParam.lblRequiredSpace.Dispatcher.Invoke({
                $ControlsParam.lblRequiredSpace.Content = $requiredSpaceString
            }, [System.Windows.Threading.DispatcherPriority]::Background)
            Write-Host "Updated Required Space Label: $requiredSpaceString"
        }

        # --- Window Initialization ---
        Write-Host "Initializing window controls based on mode."
        $controls.lblMode.Content = if ($IsBackup) { "Mode: Backup" } else { "Mode: Restore" }
        $controls.btnStart.Content = if ($IsBackup) { "Backup" } else { "Restore" }
        $controls.btnAddBAUPaths.IsEnabled = $IsBackup

        # Set default path & update free space
        Write-Host "Checking default path: $script:DefaultPath"
        if (-not (Test-Path $script:DefaultPath)) {
            Write-Host "Default path not found. Attempting to create."
            try {
                New-Item -Path $script:DefaultPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
                Write-Host "Default path created."
            } catch {
                 Write-Warning "Could not create default path: $script:DefaultPath. Please select a location manually."
                 $script:DefaultPath = $env:USERPROFILE # Fallback
                 Write-Host "Using fallback default path: $script:DefaultPath"
            }
        }
        $controls.txtSaveLoc.Text = $script:DefaultPath
        & $script:UpdateFreeSpaceLabel -ControlsParam $controls

        # Load initial items based on mode
        Write-Host "Loading initial items for ListView based on mode."
        $initialItemsList = [System.Collections.Generic.List[PSCustomObject]]::new()
        if ($IsBackup) {
            Write-Host "Backup Mode: Getting default paths using Get-BackupPaths."
            try {
                $paths = Get-BackupPaths
                Write-Host "Get-BackupPaths returned $($paths.Count) items."
                if ($paths -ne $null -and $paths.Count -gt 0) {
                    $paths | ForEach-Object { $initialItemsList.Add(($_ | Add-Member -MemberType NoteProperty -Name 'IsSelected' -Value $true -PassThru)) }
                }
            } catch { Write-Error "Error calling Get-BackupPaths: $($_.Exception.Message)" }
        }
        elseif (Test-Path $script:DefaultPath) {
            Write-Host "Restore Mode: Checking for latest backup in $script:DefaultPath."
            $latestBackup = Get-ChildItem -Path $script:DefaultPath -Directory -Filter "Backup_*" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($latestBackup) {
                Write-Host "Found latest backup: $($latestBackup.FullName)"
                $controls.txtSaveLoc.Text = $latestBackup.FullName
                & $script:UpdateFreeSpaceLabel -ControlsParam $controls
                $backupItems = Get-ChildItem -Path $latestBackup.FullName | Where-Object { $_.Name -notmatch '^(FileList_.*\.csv|Drives\.csv|Printers\.txt)$' } | ForEach-Object { [PSCustomObject]@{ Name = $_.Name; Type = if ($_.PSIsContainer) { "Folder" } else { "File" }; Path = $_.FullName; IsSelected = $true } }
                $backupItems | ForEach-Object { $initialItemsList.Add($_) }
                Write-Host "Populated ListView with $($initialItemsList.Count) items from latest backup."
            } else {
                 Write-Host "No backups found in $script:DefaultPath."
                 $controls.lblStatus.Content = "Restore mode: No backups found in $script:DefaultPath. Please browse."
            }
        } else { Write-Host "Restore Mode: Default path $script:DefaultPath does not exist." }

        if ($controls['lvwFiles'] -ne $null) {
            $controls.lvwFiles.ItemsSource = $initialItemsList
            Write-Host "Assigned $($initialItemsList.Count) initial items to ListView ItemsSource."
            & $script:UpdateRequiredSpaceLabel -ControlsParam $controls
        } else { Write-Error "ListView control ('lvwFiles') not found!"; throw "ListView control ('lvwFiles') could not be found." }
        Write-Host "Finished loading initial items."

        # --- Event Handlers ---
        Write-Host "Assigning event handlers."
        # btnBrowse Handler
        $controls.btnBrowse.Add_Click({
            Write-Host "Browse button clicked."
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $dialog.Description = if ($IsBackup) { "Select location to save backup" } else { "Select backup folder to restore from" }
            if(Test-Path $controls.txtSaveLoc.Text){ $dialog.SelectedPath = $controls.txtSaveLoc.Text } else { $dialog.SelectedPath = $script:DefaultPath }
            $dialog.ShowNewFolderButton = $IsBackup
            $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized' }
            $result = $dialog.ShowDialog($owner); $owner.Dispose()
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $selectedPath = $dialog.SelectedPath; Write-Host "Folder selected: $selectedPath"
                $controls.txtSaveLoc.Text = $selectedPath; & $script:UpdateFreeSpaceLabel -ControlsParam $controls
                if (-not $IsBackup) {
                    Write-Host "Restore Mode: Loading items from selected backup folder."
                    $logFilePath = Join-Path $selectedPath "FileList_Backup.csv"
                    $itemsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                    if (Test-Path -Path $logFilePath) {
                         Write-Host "Log file found. Populating ListView."
                         $backupItems = Get-ChildItem -Path $selectedPath | Where-Object { $_.Name -notmatch '^(FileList_.*\.csv|Drives\.csv|Printers\.txt)$' } | ForEach-Object { [PSCustomObject]@{ Name = $_.Name; Type = if ($_.PSIsContainer) { "Folder" } else { "File" }; Path = $_.FullName; IsSelected = $true } }
                         $backupItems | ForEach-Object { $itemsList.Add($_) }
                         $controls.lblStatus.Content = "Ready to restore from: $selectedPath"
                         Write-Host "ListView updated with $($itemsList.Count) items."
                    } else {
                         Write-Warning "Selected folder is not a valid backup (missing FileList_Backup.csv)."
                         $controls.lblStatus.Content = "Selected folder is not a valid backup."; [System.Windows.MessageBox]::Show("Selected folder is not a valid backup.", "Invalid Backup Folder", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    }
                    $controls.lvwFiles.ItemsSource = $itemsList; & $script:UpdateRequiredSpaceLabel -ControlsParam $controls
                } else { $controls.lblStatus.Content = "Backup location set to: $selectedPath" }
            } else { Write-Host "Folder selection cancelled."}
        })
        # btnAddFile Handler
        $controls.btnAddFile.Add_Click({
            Write-Host "Add File button clicked."
            $dialog = New-Object System.Windows.Forms.OpenFileDialog; $dialog.Title = "Select File(s) to Add"; $dialog.Multiselect = $true
            $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized' }
            $result = $dialog.ShowDialog($owner); $owner.Dispose()
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                Write-Host "$($dialog.FileNames.Count) file(s) selected."
                $currentItems = $controls.lvwFiles.ItemsSource
                $newItemsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                if ($currentItems -ne $null) { foreach ($item in $currentItems) { $newItemsList.Add($item) } }
                $existingPaths = $newItemsList.Path
                $addedCount = 0
                foreach ($file in $dialog.FileNames) {
                    if (-not ($existingPaths -contains $file)) {
                        Write-Host "Adding file: $file"
                        $newItemsList.Add([PSCustomObject]@{ Name = [System.IO.Path]::GetFileName($file); Type = "File"; Path = $file; IsSelected = $true })
                        $addedCount++
                    } else { Write-Host "Skipping duplicate file: $file"}
                }
                if ($addedCount -gt 0) { $controls.lvwFiles.ItemsSource = $newItemsList; Write-Host "Added $addedCount new file(s)."; & $script:UpdateRequiredSpaceLabel -ControlsParam $controls }
            } else { Write-Host "File selection cancelled."}
        })
        # btnAddFolder Handler
        $controls.btnAddFolder.Add_Click({
             Write-Host "Add Folder button clicked."
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog; $dialog.Description = "Select Folder to Add"; $dialog.ShowNewFolderButton = $false
            $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized' }
            $result = $dialog.ShowDialog($owner); $owner.Dispose()
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                 $selectedPath = $dialog.SelectedPath; Write-Host "Folder selected to add: $selectedPath"
                 $currentItems = $controls.lvwFiles.ItemsSource
                 $newItemsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                 if ($currentItems -ne $null) { foreach ($item in $currentItems) { $newItemsList.Add($item) } }
                 $existingPaths = $newItemsList.Path
                 if (-not ($existingPaths -contains $selectedPath)) {
                    Write-Host "Adding folder: $selectedPath"
                    $newItemsList.Add([PSCustomObject]@{ Name = [System.IO.Path]::GetFileName($selectedPath); Type = "Folder"; Path = $selectedPath; IsSelected = $true })
                    $controls.lvwFiles.ItemsSource = $newItemsList; Write-Host "Updated ListView with new folder."; & $script:UpdateRequiredSpaceLabel -ControlsParam $controls
                 } else { Write-Host "Skipping duplicate folder: $selectedPath" }
            } else { Write-Host "Folder selection cancelled."}
        })
        # btnAddBAUPaths Handler
        $controls.btnAddBAUPaths.Add_Click({
            Write-Host "Add BAU Paths button clicked."
            if (-not $IsBackup) { Write-Warning "Add BAU Paths button disabled in Restore mode."; return }
            $bauPaths = $null; try { $bauPaths = Get-UserPaths } catch { Write-Error "Error calling Get-UserPaths: $($_.Exception.Message)"; [System.Windows.MessageBox]::Show("Error retrieving user folders: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error); return }
            if ($bauPaths -ne $null -and $bauPaths.Count -gt 0) {
                $currentItems = $controls.lvwFiles.ItemsSource
                $newItemsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                if ($currentItems -ne $null) { foreach ($item in $currentItems) { $newItemsList.Add($item) } }
                $existingPaths = $newItemsList.Path
                $addedCount = 0
                foreach ($bauItem in $bauPaths) {
                    if (Test-Path -LiteralPath $bauItem.Path -PathType Container) {
                        if (-not ($existingPaths -contains $bauItem.Path)) {
                            Write-Host "Adding BAU path: $($bauItem.Path)"
                            $newItemsList.Add([PSCustomObject]@{ Name = $bauItem.Name; Type = $bauItem.Type; Path = $bauItem.Path; IsSelected = $true })
                            $addedCount++
                        } else { Write-Host "Skipping duplicate BAU path: $($bauItem.Path)" }
                    } else { Write-Host "Skipping non-existent BAU path: $($bauItem.Path)" }
                }
                if ($addedCount -gt 0) { $controls.lvwFiles.ItemsSource = $newItemsList; Write-Host "Added $addedCount new BAU path(s)."; & $script:UpdateRequiredSpaceLabel -ControlsParam $controls }
                else { Write-Host "No new BAU paths added."; [System.Windows.MessageBox]::Show("No new user folders added.", "Info", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) }
            } else { Write-Host "Get-UserPaths returned no valid paths."; [System.Windows.MessageBox]::Show("Could not find standard user folders.", "Info", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) }
        })
        # btnRemove Handler
        $controls.btnRemove.Add_Click({
            Write-Host "Remove Selected button clicked."
            $selectedObjects = @($controls.lvwFiles.SelectedItems)
            if ($selectedObjects.Count -gt 0) {
                Write-Host "Removing $($selectedObjects.Count) selected item(s)."
                $itemsToKeep = [System.Collections.Generic.List[PSCustomObject]]::new()
                if ($controls.lvwFiles.ItemsSource -ne $null) {
                    $controls.lvwFiles.ItemsSource | ForEach-Object { $currentItem = $_; $isMarkedForRemoval = $false; foreach ($itemToRemove in $selectedObjects) { if ($currentItem.Path -eq $itemToRemove.Path) { $isMarkedForRemoval = $true; break } }; if (-not $isMarkedForRemoval) { $itemsToKeep.Add($currentItem) } }
                }
                $controls.lvwFiles.ItemsSource = $itemsToKeep; Write-Host "Kept $($itemsToKeep.Count) items."; & $script:UpdateRequiredSpaceLabel -ControlsParam $controls
            } else { Write-Host "No items selected to remove."}
        })
        # lvwFiles Checkbox Click Handler
        $controls.lvwFiles.Add_PreviewMouseLeftButtonUp({
            param($sender, $e)
            if ($e.OriginalSource -is [System.Windows.Controls.CheckBox]) {
                Write-Host "Checkbox clicked within ListView."
                $sender.Dispatcher.InvokeAsync([action]{ Write-Host "Executing scheduled UpdateRequiredSpaceLabel."; & $script:UpdateRequiredSpaceLabel -ControlsParam $controls }, [System.Windows.Threading.DispatcherPriority]::ContextIdle) | Out-Null
            }
        })
        Write-Host "Assigned event handlers."

        # --- Start Button Logic (CORRECTED STRUCTURE) ---
        $controls.btnStart.Add_Click({
            $modeStringLocal = if ($IsBackup) { 'Backup' } else { 'Restore' } # Use local var
            Write-Host "Start button clicked. Mode: $modeStringLocal"

            $location = $controls.txtSaveLoc.Text
            Write-Host "Selected location: $location"
            if ([string]::IsNullOrEmpty($location) -or -not (Test-Path $location -PathType Container)) {
                [System.Windows.MessageBox]::Show("Please select a valid target directory first.", "Location Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }

            # Final check/update of required space
            Write-Host "Recalculating required space before starting..."
            & $script:UpdateRequiredSpaceLabel -ControlsParam $controls

            # Get items and options from UI
            $itemsFromUI = [System.Collections.Generic.List[PSCustomObject]]::new()
            $checkedItems = @($controls.lvwFiles.ItemsSource) | Where-Object { $_.IsSelected }
            if ($checkedItems) { $checkedItems | ForEach-Object { $itemsFromUI.Add($_) } }

            $doNetwork = $controls.chkNetwork.IsChecked
            $doPrinters = $controls.chkPrinters.IsChecked

            # Disable UI
            Write-Host "Disabling UI controls and setting wait cursor."
            $controls | ForEach-Object { if ($_.Value -is [System.Windows.Controls.Control]) { $_.Value.IsEnabled = $false } }
            $window.Cursor = [System.Windows.Input.Cursors]::Wait
            $controls.prgProgress.IsIndeterminate = $true
            $controls.prgProgress.Value = 0
            $controls.txtProgress.Text = "Initializing..."
            $controls.lblStatus.Content = "Starting $modeStringLocal..."

            # Define Progress Action for UI updates
            $uiProgressAction = {
                param($status, $percent, $details)
                # Update UI elements on the UI thread
                $window.Dispatcher.InvokeAsync(
                    [action]{
                        $controls.lblStatus.Content = $status
                        $controls.txtProgress.Text = $details
                        if ($percent -ge 0) {
                            $controls.prgProgress.IsIndeterminate = $false
                            $controls.prgProgress.Value = $percent
                        } else { # Error state
                            $controls.prgProgress.IsIndeterminate = $false
                            $controls.prgProgress.Value = 0
                        }
                    },
                    [System.Windows.Threading.DispatcherPriority]::Background
                ) | Out-Null
            }

            # Execute Core Logic
            $success = $false
            try { # INNER try block starts here
                if ($IsBackup) {
                    # Construct backup path
                    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                    $username = $env:USERNAME -replace '[^a-zA-Z0-9]', '_'
                    $backupRootPath = Join-Path $location "Backup_${username}_$timestamp"

                    if ($itemsFromUI.Count -eq 0) { throw "No items selected (checked) for backup." }

                    $success = Invoke-BackupOperation -BackupRootPath $backupRootPath `
                                                     -ItemsToBackup $itemsFromUI `
                                                     -BackupNetworkDrives $doNetwork `
                                                     -BackupPrinters $doPrinters `
                                                     -ProgressAction $uiProgressAction
                } else { # Restore
                    $backupRootPath = $location # Location IS the backup path

                    # --- Filtering Logic (Example - Needs Testing) ---
                    $csvLogPath = Join-Path $backupRootPath "FileList_Backup.csv"
                    if (-not (Test-Path $csvLogPath -PathType Leaf)) { throw "Backup log file '$csvLogPath' not found." }
                    $backupLog = Import-Csv -Path $csvLogPath -Encoding UTF8
                    if (-not $backupLog) { throw "Backup log file is empty or could not be read." }

                    $selectedTopLevelNames = $itemsFromUI | Select-Object -ExpandProperty Name
                    $logEntriesToRestore = $backupLog | Where-Object {
                        $topLevelName = ($_.BackupRelativePath -split '[\\/]', 2)[0]
                        $selectedTopLevelNames -contains $topLevelName
                    }
                    if (-not $logEntriesToRestore) { throw "None of the selected items correspond to entries in the backup log." }
                    # --- End Filtering Logic ---

                    Write-Warning "UI Restore currently triggers restore of ALL items from log, not just selected ones. Modify Invoke-RestoreOperation for selective restore."

                    $success = Invoke-RestoreOperation -BackupRootPath $backupRootPath `
                                                     -RestoreAllFromLog ` # Using this for now
                                                     -RestoreNetworkDrives $doNetwork `
                                                     -RestorePrinters $doPrinters `
                                                     -ProgressAction $uiProgressAction
                } # End of if/else for Backup/Restore

                # --- Operation Completion (UI Feedback) ---
                if ($success) {
                    Write-Host "Operation completed successfully (UI)." -ForegroundColor Green
                    $controls.lblStatus.Content = "Operation completed successfully."
                    [System.Windows.MessageBox]::Show("The $modeStringLocal operation completed successfully!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                } else {
                    Write-Error "Operation failed (UI)."
                    $controls.lblStatus.Content = "Operation Failed!"
                    [System.Windows.MessageBox]::Show("The $modeStringLocal operation failed. Check details.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                } # End of if/else for success/failure message

            # --- MOVED CATCH/FINALLY INSIDE Add_Click ---
            } catch { # INNER catch block
                # --- Catch unexpected errors during the call ---
                $errorMessage = "Operation Failed (Inner Catch): $($_.Exception.Message)"
                Write-Error $errorMessage
                $controls.lblStatus.Content = "Operation Failed!"
                $controls.txtProgress.Text = $errorMessage
                $controls.prgProgress.Value = 0
                $controls.prgProgress.IsIndeterminate = $false
                [System.Windows.MessageBox]::Show($errorMessage, "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            } finally { # INNER finally block
                Write-Host "Operation finished (inner finally block). Re-enabling UI controls."
                # Re-enable UI elements
                 $controls | ForEach-Object {
                     if ($_.Value -is [System.Windows.Controls.Control]) {
                         if ($_.Key -eq 'btnAddBAUPaths') { $_.Value.IsEnabled = $IsBackup }
                         else { $_.Value.IsEnabled = $true }
                     }
                 }
                 $window.Cursor = [System.Windows.Input.Cursors]::Arrow
                 Write-Host "Cursor reset."
            } # End of INNER finally

        }) # End btnStart.Add_Click scriptblock

        # --- Show Window ---
        Write-Host "Showing main window."
        $window.ShowDialog() | Out-Null
        Write-Host "Main window closed."

    # --- OUTER Catch/Finally for Window Load ---
    } catch { # OUTER catch block
        # --- Window Load Failure ---
        $errorMessage = "Failed to load main window: $($_.Exception.Message)"
        Write-Error $errorMessage; Write-Host $errorMessage -ForegroundColor Red
        try { [System.Windows.MessageBox]::Show($errorMessage, "Critical Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) } catch {}
    } finally { # OUTER finally block
         Write-Host "Exiting Show-MainWindow function."
         Remove-Variable -Name UpdateFreeSpaceLabel -Scope Script -ErrorAction SilentlyContinue
         Remove-Variable -Name UpdateRequiredSpaceLabel -Scope Script -ErrorAction SilentlyContinue
    } # End of OUTER finally

} # End of Show-MainWindow function

#endregion Functions

# --- Main Execution ---
Write-Host "--- Script Starting ---"
Clear-Variable -Name updateJob -Scope Script -ErrorAction SilentlyContinue

# Ensure Default Path Exists
if (-not (Test-Path $script:DefaultPath)) {
    Write-Host "Default path '$($script:DefaultPath)' not found. Attempting to create."
    try {
        New-Item -Path $script:DefaultPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
        Write-Host "Default path created."
    } catch {
         Write-Warning "Could not create default path: $($script:DefaultPath). Express mode might fail."
    }
}

try {
    # Determine mode
    Write-Host "Calling Show-ModeDialog to determine operation mode."
    $script:selectedMode = Show-ModeDialog

    # Handle mode selection
    switch ($script:selectedMode) {
        'Backup' {
            Write-Host "Mode selected: Backup. Showing main window."
            Show-MainWindow -IsBackup $true
        }
        'Restore' {
            Write-Host "Mode selected: Restore. Starting background updates and showing main window."
            Start-BackgroundUpdateJob # Start updates for Restore mode
            Show-MainWindow -IsBackup $false
        }
        'Express' {
            Write-Host "Mode selected: Express. Executing Express logic..."
            Write-Host "Checking for recent backup in '$($script:DefaultPath)'..."
            $todayDateStr = Get-Date -Format "yyyyMMdd"
            $latestBackup = Get-ChildItem -Path $script:DefaultPath -Directory -Filter "Backup_*" |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1

            $restoreCandidatePath = $null
            if ($latestBackup -and $latestBackup.Name -match "_${todayDateStr}_") {
                Write-Host "Recent backup found: $($latestBackup.FullName)" -ForegroundColor Green
                $restoreCandidatePath = $latestBackup.FullName
            } else {
                if ($latestBackup) { Write-Host "Latest backup found ($($latestBackup.Name)) is not from today." }
                else { Write-Host "No existing backups found in '$($script:DefaultPath)'." }
            }

            # Define Progress Action for Console
            $consoleProgressAction = {
                param($status, $percent, $details)
                # Simple console output
                 Write-Host "[$($status) - $percent%]: $details"
                 # Optionally add Write-Progress here
                 # Write-Progress -Activity $status -Status "$percent% Complete" -CurrentOperation $details -PercentComplete $percent
            }

            if ($restoreCandidatePath) {
                # --- Express Restore Flow ---
                Write-Host "Starting Express Restore from '$restoreCandidatePath'..." -ForegroundColor Yellow
                Start-BackgroundUpdateJob # Start updates for Restore mode

                # Determine options (assume restore all available)
                $restoreDrives = Test-Path (Join-Path $restoreCandidatePath "Drives.csv")
                $restorePrinters = Test-Path (Join-Path $restoreCandidatePath "Printers.txt")
                Write-Host "Restore Options - Drives: $restoreDrives, Printers: $restorePrinters"

                $success = Invoke-RestoreOperation -BackupRootPath $restoreCandidatePath `
                                                 -RestoreAllFromLog `
                                                 -RestoreNetworkDrives $restoreDrives `
                                                 -RestorePrinters $restorePrinters `
                                                 -ProgressAction $consoleProgressAction
                if ($success) {
                    Write-Host "Express Restore completed successfully." -ForegroundColor Green
                    [System.Windows.MessageBox]::Show("Express Restore completed successfully from `n$restoreCandidatePath", "Express Restore Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                } else {
                    Write-Error "Express Restore failed."
                    [System.Windows.MessageBox]::Show("Express Restore failed. Check console output for details.", "Express Restore Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                }

            } else {
                # --- Express Backup Flow ---
                Write-Host "Starting Express Backup to '$($script:DefaultPath)'..." -ForegroundColor Yellow

                # Determine items (Defaults + BAU)
                $itemsToBackup = [System.Collections.Generic.List[PSCustomObject]]::new()
                try { Get-BackupPaths | ForEach-Object { $itemsToBackup.Add($_) } } catch { Write-Warning "Error getting default backup paths: $($_.Exception.Message)"}
                # Davids BAU folders
                #try { Get-UserPaths | ForEach-Object { $itemsToBackup.Add($_) } } catch { Write-Warning "Error getting BAU paths: $($_.Exception.Message)"}

                # Add IsSelected property (assume all true for Express)
                $itemsToBackupFinal = $itemsToBackup | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name 'IsSelected' -Value $true -PassThru }

                # Determine options (assume backup all)
                $backupDrives = $true
                $backupPrinters = $true
                Write-Host "Backup Options - Drives: $backupDrives, Printers: $backupPrinters"
                Write-Host "Items to back up:"
                $itemsToBackupFinal | Format-Table Name, Type, Path -AutoSize

                # Construct backup path
                $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                $username = $env:USERNAME -replace '[^a-zA-Z0-9]', '_'
                $backupRootPath = Join-Path $script:DefaultPath "Backup_${username}_$timestamp"

                $success = Invoke-BackupOperation -BackupRootPath $backupRootPath `
                                                 -ItemsToBackup $itemsToBackupFinal `
                                                 -BackupNetworkDrives $backupDrives `
                                                 -BackupPrinters $backupPrinters `
                                                 -ProgressAction $consoleProgressAction
                if ($success) {
                    Write-Host "Express Backup completed successfully." -ForegroundColor Green
                    [System.Windows.MessageBox]::Show("Express Backup completed successfully to `n$backupRootPath", "Express Backup Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                } else {
                    Write-Error "Express Backup failed."
                    [System.Windows.MessageBox]::Show("Express Backup failed. Check console output for details.", "Express Backup Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                }
            }
        }
        'Cancel' {
            Write-Host "Operation cancelled by user."
        }
        Default {
            Write-Warning "Invalid mode returned from dialog: $script:selectedMode"
        }
    }

} catch {
    # Catch errors during initial mode selection or Express logic execution
    $errorMessage = "An unexpected error occurred: $($_.Exception.Message)"
    Write-Error $errorMessage; Write-Host $errorMessage -ForegroundColor Red
    try { [System.Windows.MessageBox]::Show($errorMessage, "Fatal Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) } catch {}
} finally {
    # Check for and receive output from the background job AFTER everything finishes
    if ((Get-Variable -Name updateJob -Scope Script -ErrorAction SilentlyContinue) -ne $null -and $script:updateJob -ne $null) {
        Write-Host "`n--- Waiting for background update job (GPUpdate/CM Actions) to complete... ---" -ForegroundColor Yellow
        Wait-Job $script:updateJob | Out-Null
        Write-Host "--- Background Update Job Output (GPUpdate/CM Actions): ---" -ForegroundColor Yellow
        Receive-Job $script:updateJob
        Remove-Job $script:updateJob
        Write-Host "--- End of Background Update Job Output ---" -ForegroundColor Yellow
    } else {
        Write-Host "`nNo background update job was started or it was already cleaned up." -ForegroundColor Gray
    }
}

Write-Host "--- Script Execution Finished ---"
if ($Host.Name -eq 'ConsoleHost' -and -not $psISE -and $env:TERM_PROGRAM -ne 'vscode') {
    Write-Host "Press Enter to exit..." -ForegroundColor Yellow
    Read-Host
}