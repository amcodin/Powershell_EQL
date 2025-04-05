# Requires -Version 3.0
# Everything working except remote, need testing

##############################################################################################################################
#        Client Data Backup Tool         											#
#       Written By Jared Vosters - github.com/cattboy        		#
# Based off the script by Stephen Onions & Kevin King UserBackupRefresh_Persist 1	#
#
# This Powershell script runs locally only, has 4 prompts. EXPRESS, BACKUP, RESTORE, REMOTE
#
# Express - No gui, fully automated. Does not require admin.
# Looks for a backupfolder in C:/localdata with todays date. (Folder can be created by the Backup prompt)
# If found, it will automatically restore files from backup folder to desginated locations, defined in FileList_Backup.csv
# If no backup folder is found, create backup folder with all options checked (similar to pressing Restore prompt)
#
# Backup - GUI interface. Does not require admin.
# Select folders/files to backup, backup location, etc
#
# Restore - GUI interface. Does not require admin initially.
# Select backup folder to restore.
# Executes GPUpdate /force and ConfigMgr actions locally (these might require elevation implicitly, handled by the OS/job).
#
# Remote - Console-based. REQUIRES ADMIN.
# If not run as admin, script will attempt to restart itself elevated.
# Prompts for OLD Device Asset number/IP.
# Creates map network drive to old device C$.
# Copies required files (defined by Get-BackupPaths) from old device to a new backup folder on the NEW device (C:\LocalData).
# Removes map network drive.
# Restores files locally from the newly created backup folder.
# Executes GPUpdate /force and ConfigMgr actions locally.
#
###########################################################################################################################

# --- Parameter for Elevated Re-launch ---
param(
    # Internal switch used when re-launching the script as admin for Remote mode
    [switch]$ElevatedRemoteRun
)

# --- Function to check for Admin privileges ---
function Test-IsAdmin {
    try {
        $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        Write-Warning "Could not determine administrator status: $($_.Exception.Message)"
        return $false # Assume not admin if check fails
    }
}

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
$script:RemoteMappedDriveLetter = "Z" # Drive letter to use for remote mapping
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
        [scriptblock]$ProgressAction,

        [Parameter(Mandatory=$false)]
        [string]$SettingsSourceComputer = $env:COMPUTERNAME
    )

    Write-Host "--- Starting Backup Operation to '$BackupRootPath' ---"
    $UpdateProgress = { if ($ProgressAction) { $ProgressAction.Invoke($args[0], $args[1], $args[2]) } else { Write-Host "$($args[0]) ($($args[1])%) - $($args[2])" } }
    $isRemoteSettings = $SettingsSourceComputer -ne $env:COMPUTERNAME

    # Main try block for the entire operation
    try {
        $UpdateProgress.Invoke("Initializing Backup", 0, "Creating backup directory...")
        if (-not (Test-Path $BackupRootPath)) {
            New-Item -Path $BackupRootPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Write-Host "Backup directory created: $BackupRootPath"
        }

        $csvLogPath = Join-Path $BackupRootPath "FileList_Backup.csv"
        Write-Host "Creating log file: $csvLogPath"
        # Ensure headers match expected restore columns
        "OriginalFullPath,BackupRelativePath,SourcePath" | Set-Content -Path $csvLogPath -Encoding UTF8

        if (-not $ItemsToBackup -or $ItemsToBackup.Count -eq 0) {
             throw "No items specified for backup."
        }
        Write-Host "Processing $($ItemsToBackup.Count) items for backup."

        # Estimate total files for progress
        $UpdateProgress.Invoke("Estimating size...", 0, "Calculating total files...")
        $totalFilesEstimate = 0L
        foreach ($item in $ItemsToBackup) {
            # FIX: Simplified check - rely on IsNullOrEmpty and Test-Path. Avoids '.PSObject.Properties.Contains' error.
            # Assumes Get-BackupPaths/Get-UserPaths always provide 'Path' property.
            if ((-not [string]::IsNullOrEmpty($item.Path)) -and (Test-Path -LiteralPath $item.Path -ErrorAction SilentlyContinue)) {
                if ($item.Type -eq 'Folder') {
                    try { $totalFilesEstimate += (Get-ChildItem -LiteralPath $item.Path -Recurse -File -Force -ErrorAction SilentlyContinue).Count } catch {}
                } else { $totalFilesEstimate++ }
            }
        } # End foreach for size estimation

        $networkDriveCount = $(if ($BackupNetworkDrives) { 1 } else { 0 })
        $printerCount = $(if ($BackupPrinters) { 1 } else { 0 })
        if ($BackupNetworkDrives) { $totalFilesEstimate++ }
        if ($BackupPrinters) { $totalFilesEstimate++ }
        Write-Host "Estimated total files/items: $totalFilesEstimate"

        $currentItemIndex = 0
        $totalItems = $ItemsToBackup.Count + $networkDriveCount + $printerCount

        # Process Files/Folders
        foreach ($item in $ItemsToBackup) {
            $currentItemIndex++
            $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 100) } else { 0 }
            $statusMessage = "Backing up Item $currentItemIndex of $totalItems"
            $UpdateProgress.Invoke($statusMessage, $percentComplete, "Processing: $($item.Name)")
            Write-Host "Processing item: $($item.Name) ($($item.Type)) - Source Path: $($item.Path) -> Local Target Path: $($item.LocalDestinationPath)"
            $sourcePath = $item.Path
            # Use the LocalDestinationPath provided by the item object (should be resolved by Get-BackupPaths/Get-UserPaths)
            $localDestinationPath = $item.LocalDestinationPath

            # FIX: Simplified check - rely on IsNullOrEmpty and Test-Path. Avoids '.PSObject.Properties.Contains' error.
            if ((-not [string]::IsNullOrEmpty($sourcePath)) -and (Test-Path -LiteralPath $sourcePath)) {
                # Path exists, proceed with copy logic

                # Ensure LocalDestinationPath is valid before logging/using it
                if ([string]::IsNullOrEmpty($localDestinationPath)) {
                    Write-Warning "Item '$($item.Name)' has an invalid or empty LocalDestinationPath. Using source path '$sourcePath' as fallback for logging."
                    $localDestinationPath = $sourcePath # Fallback for logging only
                }

                if ($item.Type -eq "Folder") {
                    Write-Host "  Item is a folder. Processing recursively..."
                    try {
                        $basePathLength = $sourcePath.TrimEnd('\').Length
                        # Use item name for the relative path base in the backup
                        $backupRelativeBase = $item.Name
                        $targetBackupBaseDir = Join-Path $BackupRootPath $backupRelativeBase
                        if (-not (Test-Path $targetBackupBaseDir)) {
                            New-Item -Path $targetBackupBaseDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
                        }
                        # Log the top-level folder mapping
                        "`"$localDestinationPath`",`"$backupRelativeBase`",`"$sourcePath`"" | Add-Content -Path $csvLogPath -Encoding UTF8

                        # Get all files within the source folder
                        Get-ChildItem -LiteralPath $sourcePath -Recurse -File -Force -ErrorAction Stop | ForEach-Object {
                            $originalFileFullPath = $_.FullName
                            # Calculate relative path within the source folder structure
                            $relativeFilePath = $originalFileFullPath.Substring($basePathLength).TrimStart('\')
                            # Construct the relative path within the backup structure
                            $backupRelativePath = Join-Path $backupRelativeBase $relativeFilePath
                            # Construct the full path to the file's destination within the backup
                            $targetBackupPath = Join-Path $BackupRootPath $backupRelativePath
                            # Determine the directory for the target file within the backup
                            $targetBackupDir = [System.IO.Path]::GetDirectoryName($targetBackupPath)
                            # Determine the final destination path on the local system for restore
                            $localFileDestinationPath = Join-Path $localDestinationPath $relativeFilePath

                            # Create the target directory in the backup if it doesn't exist
                            if (-not (Test-Path $targetBackupDir)) {
                                New-Item -Path $targetBackupDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
                            }
                            # Copy the file
                            Copy-Item -LiteralPath $originalFileFullPath -Destination $targetBackupPath -Force -ErrorAction Stop
                            # Log the individual file mapping
                            "`"$localFileDestinationPath`",`"$backupRelativePath`",`"$originalFileFullPath`"" | Add-Content -Path $csvLogPath -Encoding UTF8
                            $UpdateProgress.Invoke($statusMessage, $percentComplete, "Copied: $($_.Name)")
                        }
                        Write-Host "  Finished processing folder: $($item.Name)"
                    } catch {
                        Write-Warning "Error processing folder '$($item.Name)' ($sourcePath): $($_.Exception.Message)"
                        # Log the error for the folder
                        "`"$localDestinationPath`",`"ERROR_FOLDER_COPY: $($_.Exception.Message -replace '"','""')`",`"$sourcePath`"" | Add-Content -Path $csvLogPath -Encoding UTF8
                    }
                } else { # Single File
                    Write-Host "  Item is a file. Processing..."
                    try {
                       $originalFileFullPath = $sourcePath
                       # Use item name as the relative path for single files in the backup root
                       $backupRelativePath = $item.Name
                       $targetBackupPath = Join-Path $BackupRootPath $backupRelativePath
                       # Copy the file
                       Copy-Item -LiteralPath $originalFileFullPath -Destination $targetBackupPath -Force -ErrorAction Stop
                       # Log the file mapping (LocalDestinationPath should be the full file path for single files)
                       "`"$localDestinationPath`",`"$backupRelativePath`",`"$originalFileFullPath`"" | Add-Content -Path $csvLogPath -Encoding UTF8
                    } catch {
                       Write-Warning "Error processing file '$($item.Name)' ($sourcePath): $($_.Exception.Message)"
                       # Log the error for the file
                       "`"$localDestinationPath`",`"ERROR_FILE_COPY: $($_.Exception.Message -replace '"','""')`",`"$sourcePath`"" | Add-Content -Path $csvLogPath -Encoding UTF8
                    }
                }
            } else {
                # Path does not exist or Path property is missing/null
                Write-Warning "Source path not found or invalid, skipping: '$sourcePath' (Item: $($item.Name))"
                # Log skipped item - use fallback LocalDestinationPath if original was invalid
                $logDestPath = if ([string]::IsNullOrEmpty($localDestinationPath)) { $sourcePath } else { $localDestinationPath }
                "`"$logDestPath`",`"SKIPPED_SOURCE_NOT_FOUND_OR_INVALID`",`"$sourcePath`"" | Add-Content -Path $csvLogPath -Encoding UTF8
                continue # Skip to the next item
            }
        } # End foreach item

        # Backup Network Drives
        if ($BackupNetworkDrives) {
            $currentItemIndex++
            $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 100) } else { 0 }
            $statusMessage = "Backing up Item $currentItemIndex of $totalItems"
            $UpdateProgress.Invoke($statusMessage, $percentComplete, "Backing up network drives (Source: $SettingsSourceComputer)...")
            Write-Host "Processing Network Drives backup from '$SettingsSourceComputer'..."
            try {
                $driveData = $null
                if ($isRemoteSettings) {
                    Write-Host "  Getting drives remotely via Invoke-Command..."
                    $driveData = Invoke-Command -ComputerName $SettingsSourceComputer -ScriptBlock {
                        Get-WmiObject -Class Win32_MappedLogicalDisk -ErrorAction SilentlyContinue | Select-Object Name, ProviderName
                    } -ErrorAction Stop
                } else {
                    Write-Host "  Getting drives locally..."
                    $driveData = Get-WmiObject -Class Win32_MappedLogicalDisk -ErrorAction Stop | Select-Object Name, ProviderName
                }

                if ($driveData -ne $null) {
                    $driveData | Export-Csv -Path (Join-Path $BackupRootPath "Drives.csv") -NoTypeInformation -Encoding UTF8 -Force -ErrorAction Stop
                    Write-Host "Network drives backed up successfully." -ForegroundColor Green
                } else {
                    Write-Host "No network drive data retrieved from '$SettingsSourceComputer'."
                }
            } catch { Write-Warning "Failed to backup network drives from '$SettingsSourceComputer': $($_.Exception.Message)" }
        }

        # Backup Printers
        if ($BackupPrinters) {
            $currentItemIndex++
            $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 100) } else { 0 }
            $statusMessage = "Backing up Item $currentItemIndex of $totalItems"
            $UpdateProgress.Invoke($statusMessage, $percentComplete, "Backing up printers (Source: $SettingsSourceComputer)...")
            Write-Host "Processing Printers backup from '$SettingsSourceComputer'..."
            try {
                $printerData = $null
                if ($isRemoteSettings) {
                     Write-Host "  Getting printers remotely via Invoke-Command..."
                     $printerData = Invoke-Command -ComputerName $SettingsSourceComputer -ScriptBlock {
                         Get-WmiObject -Class Win32_Printer -Filter "Local = False AND Network = True" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name
                     } -ErrorAction Stop
                } else {
                     Write-Host "  Getting printers locally..."
                     $printerData = Get-WmiObject -Class Win32_Printer -Filter "Local = False AND Network = True" -ErrorAction Stop | Select-Object -ExpandProperty Name
                }

                if ($printerData -ne $null) {
                    $printerData | Set-Content -Path (Join-Path $BackupRootPath "Printers.txt") -Encoding UTF8 -Force -ErrorAction Stop
                    Write-Host "Printers backed up successfully." -ForegroundColor Green
                } else {
                     Write-Host "No network printer data retrieved from '$SettingsSourceComputer'."
                }
            } catch { Write-Warning "Failed to backup printers from '$SettingsSourceComputer': $($_.Exception.Message)" }
        }

        $UpdateProgress.Invoke("Backup Complete", 100, "Successfully backed up to: $BackupRootPath")
        Write-Host "--- Backup Operation Finished ---"
        return $true # Indicate success

    } catch {
        $errorMessage = "Backup Operation Failed: $($_.Exception.Message)"
        Write-Error $errorMessage
        # Ensure UpdateProgress is callable even if initialization failed partially
        if ($PSBoundParameters.ContainsKey('ProgressAction') -and $ProgressAction -ne $null) {
            # Use a distinct percent value like -1 to indicate error state in UI if needed
            try { $UpdateProgress.Invoke("Backup Failed", -1, $errorMessage) } catch {}
        }
        return $false # Indicate failure
    }
} # End of Invoke-BackupOperation function

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
        $backupLog = Import-Csv -Path $csvLogPath -Encoding UTF8 -ErrorAction Stop
        if (-not $backupLog) {
             throw "Backup log file is empty or could not be read: $csvLogPath"
        }
        # Check for essential columns used in restore logic
        if (-not ($backupLog[0].PSObject.Properties.Name -contains 'OriginalFullPath' -and $backupLog[0].PSObject.Properties.Name -contains 'BackupRelativePath')) {
            throw "Backup log file '$csvLogPath' is missing required columns 'OriginalFullPath' or 'BackupRelativePath'."
        }
        Write-Host "Imported $($backupLog.Count) entries from log file."

        $logEntriesToRestore = $null
        if ($RestoreAllFromLog) {
            # Filter out entries that were skipped or errored during backup
            $logEntriesToRestore = $backupLog | Where-Object { $_.BackupRelativePath -notmatch '^(ERROR_|SKIPPED_)' }
            Write-Host "RestoreAllFromLog specified: Processing $($logEntriesToRestore.Count) valid log entries."
        } else {
             # This mode isn't implemented based on the original script's UI logic (UI selects top-level items)
             # If selective restore based on UI selection was needed, logic would go here.
             throw "Selective restore based on UI selection is not implemented in this version."
        }

        if (-not $logEntriesToRestore) {
            # This could happen if the log only contained errors/skipped items
            Write-Warning "No valid log entries identified for restore."
            # Decide if this is an error or just nothing to do. Let's treat it as success with nothing done for files.
        }

        $networkDriveCount = $(if ($RestoreNetworkDrives -and (Test-Path $drivesCsvPath)) { 1 } else { 0 })
        $printerCount = $(if ($RestorePrinters -and (Test-Path $printersTxtPath)) { 1 } else { 0 })
        $totalItems = ($logEntriesToRestore | Measure-Object).Count + $networkDriveCount + $printerCount # Count actual items
        $currentItemIndex = 0
        Write-Host "Total items to restore (files/folders from log + drives + printers): $totalItems"

        if ($logEntriesToRestore) {
            Write-Host "Starting restore of files/folders..."
            foreach ($entry in $logEntriesToRestore) {
                $currentItemIndex++
                $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 100) } else { 0 }
                $statusMessage = "Restoring Item $currentItemIndex of $totalItems"

                # Target path on the local system where the item should be restored
                $targetRestorePath = $entry.OriginalFullPath
                # Relative path within the backup folder structure
                $backupRelativePath = $entry.BackupRelativePath
                # Full path to the item within the backup source directory
                $sourceBackupPath = Join-Path $BackupRootPath $backupRelativePath

                $UpdateProgress.Invoke($statusMessage, $percentComplete, "Restoring: $(Split-Path $targetRestorePath -Leaf)")

                # Check if the source item actually exists in the backup
                if (Test-Path -LiteralPath $sourceBackupPath) {
                    try {
                        # Ensure the target directory exists before copying
                        $targetRestoreDir = [System.IO.Path]::GetDirectoryName($targetRestorePath)
                        if (-not (Test-Path $targetRestoreDir)) {
                            Write-Host "  Creating target directory: $targetRestoreDir"
                            New-Item -Path $targetRestoreDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
                        }

                        $sourceItem = Get-Item -LiteralPath $sourceBackupPath
                        # Check if the source item in the backup is a directory
                        if ($sourceItem.PSIsContainer) {
                            # Source is a directory, restore its contents
                            Write-Host "  Copying folder contents from '$sourceBackupPath' to '$targetRestorePath'"
                            # Ensure the target directory itself exists
                            if (-not (Test-Path $targetRestorePath -PathType Container)) {
                                 New-Item -Path $targetRestorePath -ItemType Directory -Force -ErrorAction Stop | Out-Null
                            }
                            # Copy contents recursively
                            Copy-Item -LiteralPath (Join-Path $sourceBackupPath "*") -Destination $targetRestorePath -Recurse -Force -Container -ErrorAction Stop
                        } else {
                            # Source is a file, copy it directly
                            Write-Host "  Copying file from '$sourceBackupPath' to '$targetRestorePath'"
                            Copy-Item -LiteralPath $sourceBackupPath -Destination $targetRestorePath -Force -ErrorAction Stop
                        }
                    } catch { Write-Warning "Failed to restore '$targetRestorePath' from '$sourceBackupPath': $($_.Exception.Message)" }
                } else { Write-Warning "Source item not found in backup, skipping restore: $sourceBackupPath (Expected for: $targetRestorePath)" }
            }
            Write-Host "Finished restoring files/folders from log."
        } # End if ($logEntriesToRestore)

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
                        # Basic validation
                        if ($driveLetter -match '^[A-Z]$' -and $networkPath -match '^\\\\') {
                            if (-not (Test-Path -LiteralPath "$($driveLetter):")) {
                                try {
                                    Write-Host "  Mapping $driveLetter to $networkPath"
                                    # Persist only works reliably when run as Admin
                                    if (Test-IsAdmin) {
                                        New-PSDrive -Name $driveLetter -PSProvider FileSystem -Root $networkPath -Persist -Scope Global -ErrorAction Stop
                                    } else {
                                        Write-Warning "  Cannot use -Persist for drive mapping as script is not running elevated. Mapping non-persistently."
                                        New-PSDrive -Name $driveLetter -PSProvider FileSystem -Root $networkPath -Scope Global -ErrorAction Stop
                                    }
                                } catch { Write-Warning "  Failed to map drive $driveLetter`: $($_.Exception.Message)" }
                            } else { Write-Host "  Drive $driveLetter already exists, skipping." }
                        } else { Write-Warning "  Skipping invalid drive mapping entry: Name='$($_.Name)', Provider='$networkPath'" }
                    }
                    Write-Host "Finished processing network drive mappings."
                } catch { Write-Warning "Error processing network drive restorations: $($_.Exception.Message)" }
            } else { Write-Warning "Network drives backup file (Drives.csv) not found in '$BackupRootPath'." }
        }

        if ($RestorePrinters) {
            $currentItemIndex++
            $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 100) } else { 0 }
            $statusMessage = "Restoring Item $currentItemIndex of $totalItems"
            $UpdateProgress.Invoke($statusMessage, $percentComplete, "Restoring printers...")
            Write-Host "Processing Printers restore..."
            if (Test-Path $printersTxtPath) {
                 Write-Host "Found Printers.txt. Processing printers..."
                try {
                    # Requires COM object WScript.Network
                    $wsNet = New-Object -ComObject WScript.Network -ErrorAction Stop
                    Get-Content $printersTxtPath | ForEach-Object {
                        $printerPath = $_.Trim()
                        if (-not ([string]::IsNullOrWhiteSpace($printerPath)) -and $printerPath -match '^\\\\') {
                            Write-Host "  Attempting to add printer: $printerPath"
                            try {
                                 # This method adds the connection; it doesn't check if it exists first.
                                 # Re-adding usually doesn't cause issues.
                                 $wsNet.AddWindowsPrinterConnection($printerPath)
                                 Write-Host "    Added printer connection (or it already existed)."
                            } catch { Write-Warning "    Failed to add printer '$printerPath': $($_.Exception.Message)" }
                        } else { Write-Warning "  Skipping invalid or empty line in Printers.txt: '$_'" }
                    }
                    Write-Host "Finished processing printers."
                } catch { Write-Warning "Error processing printer restorations (Ensure WScript.Network COM object is available): $($_.Exception.Message)" }
            } else { Write-Warning "Printers backup file (Printers.txt) not found in '$BackupRootPath'." }
        }

        $UpdateProgress.Invoke("Restore Complete", 100, "Successfully restored from: $BackupRootPath")
        Write-Host "--- Restore Operation Finished ---"
        return $true # Indicate success

    } catch {
        $errorMessage = "Restore Operation Failed: $($_.Exception.Message)"
        Write-Error $errorMessage
        try { $UpdateProgress.Invoke("Restore Failed", -1, $errorMessage) } catch {} # Use -1 percent for error state
        return $false # Indicate failure
    }
}

# --- Background Job Starter ---
function Start-BackgroundUpdateJob {
    if (-not (Test-IsAdmin)) {
        Write-Warning "Background update job (GPUpdate/ConfigMgr) may require administrator privileges to run correctly."
    }

    Write-Host "Initiating background system updates job..." -ForegroundColor Yellow
    $script:updateJob = Start-Job -Name "BackgroundUpdates" -ScriptBlock {

        function Set-GPupdate {
            Write-Host "JOB: Initiating Group Policy update..." -ForegroundColor Cyan
            try {
                $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c gpupdate /force" -WorkingDirectory 'C:\Windows\System32' -PassThru -Wait -ErrorAction Stop
                if ($process.ExitCode -eq 0) {
                    Write-Host "JOB: Group Policy update completed successfully." -ForegroundColor Green
                } else {
                    Write-Warning "JOB: Group Policy update process finished with exit code: $($process.ExitCode)."
                }
            } catch {
                Write-Error "JOB: Failed to start GPUpdate process: $($_.Exception.Message)"
            }
            Write-Host "JOB: Exiting Set-GPupdate function."
        }

        function Start-ConfigManagerActions {
             param()
             Write-Host "JOB: Entering Start-ConfigManagerActions function."
             $ccmExecPath = "C:\Windows\CCM\ccmexec.exe"
             $wmiNamespace = "root\ccm"
             $wmiClassName = "SMS_Client"
             $scheduleMethodName = "TriggerSchedule"
             $overallSuccess = $false
             $cimAttemptedAndSucceeded = $false
             # Common ConfigMgr action schedule IDs
             $scheduleActions = @(
                 @{ID = '{00000000-0000-0000-0000-000000000021}'; Name = 'Machine Policy Retrieval & Evaluation Cycle'},
                 @{ID = '{00000000-0000-0000-0000-000000000022}'; Name = 'User Policy Retrieval & Evaluation Cycle'},
                 @{ID = '{00000000-0000-0000-0000-000000000001}'; Name = 'Hardware Inventory Cycle'},
                 @{ID = '{00000000-0000-0000-0000-000000000002}'; Name = 'Software Inventory Cycle'},
                 @{ID = '{00000000-0000-0000-0000-000000000113}'; Name = 'Software Updates Scan Cycle'},
                 # @{ID = '{00000000-0000-0000-0000-000000000101}'; Name = 'Hardware Inventory Collection Cycle'}, # Often same as 0001
                 @{ID = '{00000000-0000-0000-0000-000000000108}'; Name = 'Software Updates Assignments Evaluation Cycle'}
                 # @{ID = '{00000000-0000-0000-0000-000000000102}'; Name = 'Software Inventory Collection Cycle'} # Often same as 0002
             )
             Write-Host "JOB: Defined $($scheduleActions.Count) CM actions to trigger."
             $ccmService = Get-Service -Name CcmExec -ErrorAction SilentlyContinue
             if (-not $ccmService) { Write-Warning "JOB: CM service (CcmExec) not found. Skipping."; return $false }
             elseif ($ccmService.Status -ne 'Running') { Write-Warning "JOB: CM service (CcmExec) is not running (Status: $($ccmService.Status)). Skipping."; return $false }
             else { Write-Host "JOB: CM service (CcmExec) found and running." }

             Write-Host "JOB: Attempting Method 1: Triggering via CIM ($wmiNamespace -> $wmiClassName)..."
             $cimMethodSuccess = $true
             try {
                 # Check if the WMI class exists first
                 if (Get-CimClass -Namespace $wmiNamespace -ClassName $wmiClassName -ErrorAction SilentlyContinue) {
                     Write-Host "JOB: Attempting to invoke '$scheduleMethodName' on '$wmiClassName' in '$wmiNamespace'."
                     foreach ($action in $scheduleActions) {
                        Write-Host "JOB:   Triggering $($action.Name) (ID: $($action.ID)) via CIM."
                        try {
                            Invoke-CimMethod -Namespace $wmiNamespace -ClassName $wmiClassName -MethodName $scheduleMethodName -Arguments @{sScheduleID = $action.ID} -ErrorAction Stop
                            Write-Host "JOB:     $($action.Name) triggered successfully via CIM."
                        } catch {
                            Write-Warning "JOB:     Failed to trigger $($action.Name) via CIM: $($_.Exception.Message)"
                            $cimMethodSuccess = $false
                        }
                     }
                     if ($cimMethodSuccess) {
                         $cimAttemptedAndSucceeded = $true
                         $overallSuccess = $true
                         Write-Host "JOB: All actions successfully triggered via CIM." -ForegroundColor Green
                     } else { Write-Warning "JOB: One or more actions failed to trigger via CIM." }
                 } else {
                     Write-Warning "JOB: CIM Class '$wmiClassName' not found in namespace '$wmiNamespace'. Skipping CIM method."
                     $cimMethodSuccess = $false
                 }
             } catch {
                 Write-Error "JOB: An unexpected error occurred during CIM attempt: $($_.Exception.Message)"
                 $cimMethodSuccess = $false
             }

             # Fallback if CIM didn't work or wasn't available
             if (-not $cimAttemptedAndSucceeded) {
                 Write-Host "JOB: CIM method did not complete successfully or was not available. Attempting Method 2: Fallback via ccmexec.exe..."
                 if (Test-Path -Path $ccmExecPath -PathType Leaf) {
                     Write-Host "JOB: Found $ccmExecPath."
                     $execMethodSuccess = $true
                     foreach ($action in $scheduleActions) {
                         Write-Host "JOB:   Triggering $($action.Name) (ID: $($action.ID)) via ccmexec.exe."
                         try {
                             # Note: ccmexec trigger often requires elevation
                             $process = Start-Process -FilePath $ccmExecPath -ArgumentList "/TriggerSchedule $($action.ID)" -NoNewWindow -PassThru -Wait -ErrorAction Stop
                             if ($process.ExitCode -ne 0) {
                                 Write-Warning "JOB:     $($action.Name) via ccmexec.exe finished with exit code $($process.ExitCode). (May require elevation)"
                             } else { Write-Host "JOB:     $($action.Name) triggered via ccmexec.exe (Exit Code 0)." }
                         } catch {
                             Write-Warning "JOB:     Failed to execute ccmexec.exe for $($action.Name): $($_.Exception.Message)"
                             $execMethodSuccess = $false
                         }
                     }
                     if ($execMethodSuccess) {
                         # We can't be certain it worked without elevation, but we tried.
                         $overallSuccess = $true # Mark as attempted
                         Write-Host "JOB: Finished attempting actions via ccmexec.exe." -ForegroundColor Green
                     } else { Write-Warning "JOB: One or more actions failed to execute via ccmexec.exe." }
                 } else { Write-Warning "JOB: Fallback executable not found at $ccmExecPath." }
             }

             if ($overallSuccess) { Write-Host "JOB: CM actions attempt finished. Check ConfigMgr logs for actual execution status." -ForegroundColor Green }
             else { Write-Warning "JOB: CM actions attempt finished, but neither CIM nor ccmexec.exe methods could be confirmed as successful." }
             Write-Host "JOB: Exiting Start-ConfigManagerActions function."
             return $overallSuccess
        }

        # Execute the functions within the job
        Set-GPupdate
        Start-ConfigManagerActions
        Write-Host "JOB: Background updates finished."

    } # End of Start-Job ScriptBlock

    Write-Host "Background update job started (ID: $($script:updateJob.Id)). Output will be shown later." -ForegroundColor Yellow
}


# --- Data Gathering Functions ---
function Get-BackupPaths {
    [CmdletBinding()]
    param (
        # Allow specifying a different base profile path, e.g., for remote operations
        [string]$UserProfilePath = $env:USERPROFILE
    )

    Write-Host "Gathering backup paths based on profile: '$UserProfilePath'"
    # Define path templates relative to the UserProfilePath
    $pathTemplates = @(
        @{ SourceRelative = 'AppData\Roaming\Microsoft\Signatures'; Type = 'Folder' }
        @{ SourceRelative = 'AppData\Roaming\Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms'; Type = 'File' }
        @{ SourceRelative = 'AppData\Roaming\Microsoft\Sticky Notes\StickyNotes.snt'; Type = 'File' } # Legacy Sticky Notes
        @{ SourceRelative = 'AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\plum.sqlite'; Type = 'File' } # Modern Sticky Notes
        @{ SourceRelative = 'AppData\Roaming\google\googleearth\myplaces.kml'; Type = 'File' }
        @{ SourceRelative = 'AppData\Local\Google\Chrome\User Data\Default\Bookmarks'; Type = 'File' }
        @{ SourceRelative = 'AppData\Local\Microsoft\Edge\User Data\Default\Bookmarks'; Type = 'File' }
        # Add more paths here as needed
    )
    $result = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($templateInfo in $pathTemplates) {
        # Construct the full source path based on the provided UserProfilePath
        $sourceFullPath = Join-Path $UserProfilePath $templateInfo.SourceRelative
        $resolvedLocalDestinationPath = $null
        $itemType = $templateInfo.Type

        # Determine the equivalent *local* destination path for restore purposes
        try {
            $localEquivalentRelative = $templateInfo.SourceRelative
            # Use environment variables for the *current* machine to resolve local paths
            if ($localEquivalentRelative.StartsWith('AppData\Roaming')) {
                $localEquivalentFullPath = Join-Path $env:APPDATA ($localEquivalentRelative -replace '^AppData\\Roaming\\?', '')
            } elseif ($localEquivalentRelative.StartsWith('AppData\Local')) {
                 $localEquivalentFullPath = Join-Path $env:LOCALAPPDATA ($localEquivalentRelative -replace '^AppData\\Local\\?', '')
            } else {
                 # Assume relative to user profile if not AppData
                 $localEquivalentFullPath = Join-Path $env:USERPROFILE $templateInfo.SourceRelative
            }
            # ExpandString can resolve variables within the path if needed, though less common here
            $resolvedLocalDestinationPath = $ExecutionContext.InvokeCommand.ExpandString($localEquivalentFullPath)
        } catch {
            Write-Warning "Could not resolve local destination equivalent for '$($templateInfo.SourceRelative)'. Using source path as placeholder."
            # Fallback: Use the source path structure, assuming it might be restored to a similar relative location
            $resolvedLocalDestinationPath = Join-Path $env:USERPROFILE $templateInfo.SourceRelative
        }

        Write-Host "Checking source path: '$sourceFullPath'"
        # Check if the *source* path exists (on local or remote machine)
        if (Test-Path -LiteralPath $sourceFullPath) {
            $item = Get-Item -LiteralPath $sourceFullPath -ErrorAction SilentlyContinue
            if ($item) {
                # Create the object with all necessary properties
                $result.Add([PSCustomObject]@{
                    Name = $item.Name
                    Path = $item.FullName # Full source path (local or remote)
                    Type = if ($item.PSIsContainer) { "Folder" } else { "File" } # Verify type
                    IsSelected = $true # Default to selected for backup/restore list
                    LocalDestinationPath = $resolvedLocalDestinationPath # Where it should go on the *local* machine during restore
                })
                Write-Host "  Found: '$($item.FullName)' (Type: $($result[-1].Type)) -> Local Target: '$resolvedLocalDestinationPath'"
            }
            else { Write-Host "  Path exists but Get-Item failed: $sourceFullPath" }
        } else { Write-Host "  Path not found or resolution failed: $sourceFullPath" }
    }
    Write-Host "Finished gathering backup paths. Found $($result.Count) items."
    return $result
}


function Get-UserPaths {
    [CmdletBinding()]
    param ()
    # Define standard user folders to potentially back up
    $specificPaths = @(
        "$env:USERPROFILE\Downloads",
        "$env:USERPROFILE\Pictures",
        "$env:USERPROFILE\Videos"
        )
    $result = [System.Collections.Generic.List[PSCustomObject]]::new()
    Write-Host "Gathering standard user folder paths..."
    foreach ($path in $specificPaths) {
        $resolvedPath = try { $ExecutionContext.InvokeCommand.ExpandString($path) } catch { $null }
        # Check if the path exists and is a directory on the *local* machine
        if ($resolvedPath -and (Test-Path -LiteralPath $resolvedPath -PathType Container)) {
            $item = Get-Item -LiteralPath $resolvedPath -ErrorAction SilentlyContinue
             if ($item) {
                 # Create the object with all necessary properties
                 $result.Add([PSCustomObject]@{
                     Name = $item.Name
                     Path = $item.FullName # Source path is the local path
                     Type = "Folder"
                     IsSelected = $true # Default to selected
                     LocalDestinationPath = $item.FullName # Destination is the same as source for these standard folders
                 })
                 Write-Host "  Found User Folder: '$($item.FullName)'"
             }
             else { Write-Host "User folder path resolved but Get-Item failed: $resolvedPath" }
        } else { Write-Host "User folder path not found or not a folder: $path" }
    }
    Write-Host "Finished gathering standard user folders. Found $($result.Count) items."
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
    $form.Size = New-Object System.Drawing.Size(450, 150)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false; $form.MinimizeBox = $false; $form.HelpButton = $false

    # Explicitly cast layout variables to [int]
    [int]$buttonWidth = 80
    [int]$buttonHeight = 30
    [int]$spacing = 20
    [int]$startX = 20
    [int]$startY = 40

    try {
        $btnBackup = New-Object System.Windows.Forms.Button
        $btnBackup.Location = New-Object System.Drawing.Point($startX, $startY)
        $btnBackup.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
        $btnBackup.Text = "Backup"
        $btnBackup.DialogResult = [System.Windows.Forms.DialogResult]::Yes # Corresponds to 'Backup'
        $form.Controls.Add($btnBackup)

        $btnRestore = New-Object System.Windows.Forms.Button
        $btnRestore.Location = New-Object System.Drawing.Point(($startX + $buttonWidth + $spacing), $startY)
        $btnRestore.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
        $btnRestore.Text = "Restore"
        $btnRestore.DialogResult = [System.Windows.Forms.DialogResult]::No # Corresponds to 'Restore'
        $form.Controls.Add($btnRestore)

        $btnExpress = New-Object System.Windows.Forms.Button
        $btnExpress.Location = New-Object System.Drawing.Point(($startX + (2 * ($buttonWidth + $spacing))), $startY)
        $btnExpress.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
        $btnExpress.Text = "Express"
        $btnExpress.DialogResult = [System.Windows.Forms.DialogResult]::OK # Corresponds to 'Express'
        $form.Controls.Add($btnExpress)

        $btnRemote = New-Object System.Windows.Forms.Button
        $btnRemote.Location = New-Object System.Drawing.Point(($startX + (3 * ($buttonWidth + $spacing))), $startY)
        $btnRemote.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
        $btnRemote.Text = "Remote"
        $btnRemote.DialogResult = [System.Windows.Forms.DialogResult]::Retry # Corresponds to 'Remote'
        $form.Controls.Add($btnRemote)

        # Add a label for clarity
        $lblPrompt = New-Object System.Windows.Forms.Label
        $lblPrompt.Text = "Choose the operation mode:"
        $lblPrompt.Location = New-Object System.Drawing.Point($startX, 15)
        $lblPrompt.AutoSize = $true
        $form.Controls.Add($lblPrompt)

        # Set AcceptButton and CancelButton for better usability
        $form.AcceptButton = $btnBackup # Default action on Enter could be Backup
        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel # Standard Cancel
        $form.CancelButton = $btnCancel # Allows Esc key to cancel

    } catch {
        Write-Error "Error creating buttons in Show-ModeDialog: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error creating dialog buttons: $($_.Exception.Message)", "Dialog Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return 'Cancel' # Return 'Cancel' on error
    }

    Write-Host "Showing mode selection dialog."
    # Use a hidden owner window to prevent the dialog from going behind other windows easily
    $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized'; Opacity = 0 }; $owner.Show()
    $result = $form.ShowDialog($owner)
    $owner.Close(); $owner.Dispose()
    $form.Dispose() # Dispose the form resources
    Write-Host "Mode selection dialog closed with result: $result"

    # Map DialogResult to mode string
    $selectedMode = switch ($result) {
        ([System.Windows.Forms.DialogResult]::Yes)   { 'Backup' }
        ([System.Windows.Forms.DialogResult]::No)    { 'Restore' }
        ([System.Windows.Forms.DialogResult]::OK)    { 'Express' }
        ([System.Windows.Forms.DialogResult]::Retry) { 'Remote' }
        Default { 'Cancel' } # Handle Cancel or closing the dialog
    }

    Write-Host "Determined mode: $selectedMode" -ForegroundColor Cyan
    return $selectedMode
}

# Show main window
function Show-MainWindow {
    param(
        [Parameter(Mandatory=$true)]
        [bool]$IsBackup # True for Backup mode, False for Restore mode
    )
    $modeString = if ($IsBackup) { 'Backup' } else { 'Restore' }
    Write-Host "Entering Show-MainWindow function. Mode: $modeString"

    # Define XAML for the main window UI
    [xml]$XAML = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="User Data Backup/Restore Tool"
    Width="800"
    Height="650"
    MinWidth="600"
    MinHeight="500"
    WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/> <!-- Location Row -->
            <RowDefinition Height="Auto"/> <!-- Space Info Row -->
            <RowDefinition Height="Auto"/> <!-- List Header Row -->
            <RowDefinition Height="*"/>    <!-- List View Row -->
            <RowDefinition Height="Auto"/> <!-- Progress Text Row -->
            <RowDefinition Height="Auto"/> <!-- Progress Bar Row -->
            <RowDefinition Height="Auto"/> <!-- Button/Status Row -->
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/> <!-- Main Area -->
            <ColumnDefinition Width="Auto"/> <!-- Button Stack -->
        </Grid.ColumnDefinitions>

        <!-- Row 0: Location -->
        <Label Content="_Location:" Target="{Binding ElementName=txtSaveLoc}" Grid.Row="0" Grid.Column="0" VerticalAlignment="Center" Margin="0,0,5,0"/>
        <TextBox Name="txtSaveLoc" Grid.Row="0" Grid.Column="0" Margin="60,0,70,5" VerticalAlignment="Center" Height="24" IsReadOnly="True"/>
        <Button Name="btnBrowse" Content="_Browse..." Grid.Row="0" Grid.Column="0" HorizontalAlignment="Right" VerticalAlignment="Center" Width="65" Height="24" Margin="0,0,0,5"/>

        <!-- Row 1: Space Info -->
        <StackPanel Grid.Row="1" Grid.Column="0" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,0,0,10">
             <Label Name="lblMode" Content="Mode: -" FontWeight="Bold" Margin="0,0,20,0"/>
             <Label Name="lblFreeSpace" Content="Free Space: -" Margin="0,0,20,0"/>
             <Label Name="lblRequiredSpace" Content="Required Space: -"/>
        </StackPanel>

        <!-- Row 2: List Header -->
        <Label Content="_Files/Folders to Process:" Target="{Binding ElementName=lvwFiles}" Grid.Row="2" Grid.Column="0" FontWeight="SemiBold" Margin="0,0,0,2"/>

        <!-- Row 3: List View and Side Buttons -->
        <ListView Name="lvwFiles" Grid.Row="3" Grid.Column="0" Margin="0,0,10,0" SelectionMode="Extended">
             <ListView.View>
                <GridView>
                    <GridViewColumn Width="30">
                        <GridViewColumn.Header>
                            <!-- Optional: Add a header checkbox to select/deselect all -->
                            <!-- <CheckBox Name="chkSelectAll" /> -->
                        </GridViewColumn.Header>
                        <GridViewColumn.CellTemplate>
                            <!-- Bind CheckBox to the IsSelected property added by Get-BackupPaths/Get-UserPaths -->
                            <DataTemplate> <CheckBox IsChecked="{Binding IsSelected, Mode=TwoWay}" VerticalAlignment="Center"/> </DataTemplate>
                        </GridViewColumn.CellTemplate>
                    </GridViewColumn>
                    <GridViewColumn Header="Name" DisplayMemberBinding="{Binding Name}" Width="180"/>
                    <GridViewColumn Header="Type" DisplayMemberBinding="{Binding Type}" Width="70"/>
                    <GridViewColumn Header="Path" DisplayMemberBinding="{Binding Path}" Width="280"/>
                    <!-- Optional: Add LocalDestinationPath for debugging/info -->
                    <!-- <GridViewColumn Header="Target Path" DisplayMemberBinding="{Binding LocalDestinationPath}" Width="280"/> -->
                </GridView>
            </ListView.View>
        </ListView>

        <StackPanel Grid.Row="3" Grid.Column="1" Margin="5,0,0,0" VerticalAlignment="Top">
            <Button Name="btnAddFile" Content="Add _File..." Width="120" Height="26" Margin="0,0,0,8"/>
            <Button Name="btnAddFolder" Content="Add F_older..." Width="120" Height="26" Margin="0,0,0,8"/>
            <Button Name="btnAddBAUPaths" Content="Add _User Folders" Width="120" Height="26" Margin="0,0,0,8" ToolTip="Adds standard user folders like Documents, Desktop, Pictures etc."/>
            <Button Name="btnRemove" Content="_Remove Selected" Width="120" Height="26" Margin="0,0,0,15"/>
            <CheckBox Name="chkNetwork" Content="Network Dri_ves" IsChecked="True" Margin="5,0,0,5"/>
            <CheckBox Name="chkPrinters" Content="_Printers" IsChecked="True" Margin="5,0,0,5"/>
        </StackPanel>

        <!-- Row 4: Progress Text -->
        <TextBlock Name="txtProgress" Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="2" Text="" Margin="0,5,0,2" TextWrapping="Wrap" VerticalAlignment="Center"/>

        <!-- Row 5: Progress Bar -->
        <ProgressBar Name="prgProgress" Grid.Row="5" Grid.Column="0" Grid.ColumnSpan="2" Height="18" Margin="0,0,0,5" VerticalAlignment="Center"/>

        <!-- Row 6: Start Button and Status Label -->
        <Button Name="btnStart" Content="_Start" Grid.Row="6" Grid.Column="0" Width="100" Height="30" HorizontalAlignment="Left" IsDefault="True"/>
        <Label Name="lblStatus" Content="Ready" Grid.Row="6" Grid.Column="0" Grid.ColumnSpan="2" Margin="110,0,0,0" VerticalAlignment="Center" FontStyle="Italic"/>

    </Grid>
</Window>
'@

    try { # OUTER try for window loading and setup
        Write-Host "Parsing XAML for main window."
        $reader = New-Object System.Xml.XmlNodeReader $XAML
        $window = [Windows.Markup.XamlReader]::Load($reader)
        Write-Host "XAML loaded successfully."
        # Set DataContext if needed for complex bindings (not used much here)
        # $window.DataContext = [PSCustomObject]@{ IsRestoreMode = (-not $IsBackup) }

        Write-Host "Finding controls in main window."
        $controls = @{}
        @('txtSaveLoc', 'btnBrowse', 'btnStart', 'lblMode', 'lblStatus', 'lvwFiles',
          'btnAddFile', 'btnAddFolder', 'btnRemove', 'chkNetwork', 'chkPrinters',
          'prgProgress', 'txtProgress', 'btnAddBAUPaths', 'lblFreeSpace', 'lblRequiredSpace'
          # Add 'chkSelectAll' if header checkbox is implemented
        ) | ForEach-Object {
            $control = $window.FindName($_)
            if ($control -eq $null) { Write-Warning "Control '$_' not found in XAML!" }
            $controls[$_] = $control
        }
        Write-Host "Controls found and stored in hashtable."

        # --- Helper Functions for UI Updates (Defined within scope for encapsulation) ---
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
                         # Attempt to get UNC path space info - simplified for reliability
                         try {
                             $uri = [System.Uri]"file://$location"
                             if ($uri.Host -and $uri.Host -ne "localhost" -and $uri.Host -ne ".") {
                                 # Basic check if system accessible via ping
                                 if (Test-Connection -ComputerName $uri.Host -Count 1 -Quiet) {
                                     # Try to get share info first (most accurate for the specific share)
                                     $shareName = $uri.Segments[1].TrimEnd('/')
                                     $diskDriveLetter = $null
                                     # Add simple error handling if we can't get UNC space
                                     try {
                                         $share = Get-WmiObject -Class Win32_Share -ComputerName $uri.Host -Filter "Name='$shareName'" -ErrorAction Stop
                                         if ($share) {
                                             $diskDriveLetter = $share.Path.Substring(0, 2)
                                             $driveInfo = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$diskDriveLetter'" -ComputerName $uri.Host -ErrorAction Stop
                                             if ($driveInfo -and $driveInfo.FreeSpace -ne $null) {
                                                 $freeSpaceString = "Free Space: $(Format-Bytes $driveInfo.FreeSpace) (on $($uri.Host))"
                                             }
                                         }
                                     } catch {
                                         # Just silently use the default "N/A (UNC)" message
                                     }
                                 }
                             }
                         } catch {
                             # Just use the default "N/A (UNC)" message
                         }
                    }

                    # Handle local drive letter if found
                    if ($driveLetter) {
                        $driveInfo = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$driveLetter'" -ErrorAction SilentlyContinue
                        if ($driveInfo -and $driveInfo.FreeSpace -ne $null) {
                            $freeSpaceString = "Free Space: $(Format-Bytes $driveInfo.FreeSpace)"
                        } else { Write-Warning "Could not get free space for drive $driveLetter" }
                    }
                } catch { Write-Warning "Error getting free space for '$location': $($_.Exception.Message)" }
            }
            # SIMPLIFIED: Direct update of the label content
            $ControlsParam.lblFreeSpace.Content = $freeSpaceString
            Write-Host "Updated Free Space Label: $freeSpaceString"
        } # End UpdateFreeSpaceLabel

        $script:UpdateRequiredSpaceLabel = {
            param($ControlsParam)
            $totalSize = 0L
            $requiredSpaceString = "Required Space: Calculating..."
            # Update UI immediately to show calculation started
            $ControlsParam.lblRequiredSpace.Content = $requiredSpaceString

            # Perform calculation
            $items = $ControlsParam.lvwFiles.ItemsSource
            if ($items -is [System.Collections.IEnumerable] -and $items -ne $null) {
                # Create a temporary list to iterate over
                $itemsToCheck = [System.Collections.Generic.List[object]]::new()
                try {
                    # Ensure ItemsSource is not null before iterating
                    if ($items -ne $null) {
                        foreach($i in $items) { $itemsToCheck.Add($i) }
                    }
                } catch { Write-Warning "Error creating temporary list from ItemsSource during space calculation: $($_.Exception.Message)" }

                # Filter for items that are actually selected (checked)
                # Ensure IsSelected property exists before filtering
                $checkedItems = $itemsToCheck | Where-Object { $_.PSObject.Properties['IsSelected'] -ne $null -and $_.IsSelected }
                Write-Host "Calculating required space for $($checkedItems.Count) checked items..."

                foreach ($item in $checkedItems) {
                    # Check if Path property exists and is valid before testing
                    if ($item -is [PSCustomObject] -and $item.PSObject.Properties['Path'] -ne $null -and -not [string]::IsNullOrEmpty($item.Path)) {
                        if (Test-Path -LiteralPath $item.Path -ErrorAction SilentlyContinue) {
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
                    } else { Write-Warning "Checked item is missing 'Path' property, is not a PSCustomObject, or Path is empty." }
                }
                $requiredSpaceString = "Required Space: $(Format-Bytes $totalSize)"
            } else {
                $requiredSpaceString = "Required Space: 0 B"
            }

            # SIMPLIFIED: Direct update of the label content
            $ControlsParam.lblRequiredSpace.Content = $requiredSpaceString
            Write-Host "Updated Required Space Label: $requiredSpaceString"
        } # End UpdateRequiredSpaceLabel

        # --- Window Initialization ---
        Write-Host "Initializing window controls based on mode."
        $controls.lblMode.Content = "Mode: $modeString"
        $controls.btnStart.Content = if ($IsBackup) { "_Backup" } else { "_Restore" } # Use underscore for Alt-key shortcut
        # Disable irrelevant buttons based on mode
        $controls.btnAddFile.IsEnabled = $IsBackup
        $controls.btnAddFolder.IsEnabled = $IsBackup
        $controls.btnAddBAUPaths.IsEnabled = $IsBackup
        # Checkboxes are less relevant for Restore decision logic, but keep enabled for visual consistency
        # $controls.chkNetwork.IsEnabled = $IsBackup
        # $controls.chkPrinters.IsEnabled = $IsBackup

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
        # Initial free space calculation
        & $script:UpdateFreeSpaceLabel -ControlsParam $controls

        Write-Host "Loading initial items for ListView based on mode."
        $initialItemsList = [System.Collections.Generic.List[PSCustomObject]]::new()
        if ($IsBackup) {
            Write-Host "Backup Mode: Getting default paths using Get-BackupPaths (local)."
            try {
                # Get-BackupPaths now returns objects with IsSelected and LocalDestinationPath
                $paths = Get-BackupPaths -UserProfilePath $env:USERPROFILE
                Write-Host "Get-BackupPaths returned $($paths.Count) items."
                if ($paths -ne $null -and $paths.Count -gt 0) {
                    # FIX: Use iterative Add instead of AddRange to avoid type conversion issues
                    foreach ($item in $paths) {
                        $initialItemsList.Add($item)
                    }
                }
            } catch { Write-Error "Error calling Get-BackupPaths: $($_.Exception.Message)"; throw "Error calling Get-BackupPaths: $($_.Exception.Message)" } # Re-throw to be caught by outer handler
        }
        elseif (Test-Path $script:DefaultPath) { # Restore Mode: Look for backups
            Write-Host "Restore Mode: Checking for latest backup in $script:DefaultPath."
            $latestBackup = Get-ChildItem -Path $script:DefaultPath -Directory -Filter "Backup_*" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($latestBackup) {
                Write-Host "Found latest backup: $($latestBackup.FullName)"
                $controls.txtSaveLoc.Text = $latestBackup.FullName
                & $script:UpdateFreeSpaceLabel -ControlsParam $controls # Update free space for selected backup location (less relevant for restore)
                $logFilePath = Join-Path $latestBackup.FullName "FileList_Backup.csv"
                if (Test-Path $logFilePath) {
                    try {
                        Write-Host "Reading backup log: $logFilePath"
                        $logContent = Import-Csv $logFilePath
                        # Group by the top-level item in the backup to populate the list view
                        $logContent | Where-Object { $_.BackupRelativePath -notmatch '^(ERROR_|SKIPPED_)' } | Group-Object { ($_.BackupRelativePath -split '[\\/]', 2)[0] } | ForEach-Object {
                            $topLevelItemName = $_.Name
                            $sourcePathInBackup = Join-Path $latestBackup.FullName $topLevelItemName
                            # Determine if it was originally a file or folder based on backup content
                            $itemType = "File"; if (Test-Path $sourcePathInBackup -PathType Container -ErrorAction SilentlyContinue) { $itemType = "Folder" }
                            # Create object for ListView - Path here is the path *within the backup*
                            $initialItemsList.Add([PSCustomObject]@{
                                Name = $topLevelItemName
                                Type = $itemType
                                Path = $sourcePathInBackup # Path in the backup folder
                                IsSelected = $true # Default to selected for restore
                                LocalDestinationPath = ($_.Group | Select-Object -First 1).OriginalFullPath # Get original path from first entry in group
                            })
                        }
                        Write-Host "Populated ListView with $($initialItemsList.Count) items from backup log."
                    } catch {
                        Write-Warning "Error reading backup log '$logFilePath': $($_.Exception.Message). Cannot populate list from log."
                        $controls.lblStatus.Content = "Error reading backup log. Cannot list items."
                        # Optionally, could list directory contents as a fallback, but log is preferred
                    }
                } else {
                    Write-Warning "Backup log file not found in '$($latestBackup.FullName)'. Cannot list items."
                    $controls.lblStatus.Content = "Backup log missing. Cannot list items."
                }
            } else {
                 Write-Host "No backups found in $script:DefaultPath."
                 $controls.lblStatus.Content = "Restore mode: No backups found in $script:DefaultPath. Please browse."
            }
        } else { Write-Host "Restore Mode: Default path $script:DefaultPath does not exist." }

        # Populate ListView
        if ($controls['lvwFiles'] -ne $null) {
            # Use an ObservableCollection for better UI updates if adding/removing frequently,
            # but List works fine for initial load and less frequent updates.
            $controls.lvwFiles.ItemsSource = $initialItemsList
            Write-Host "Assigned $($initialItemsList.Count) initial items to ListView ItemsSource."
            # *** FIX: Calculate space required initially for BOTH modes ***
            & $script:UpdateRequiredSpaceLabel -ControlsParam $controls
        } else { Write-Error "ListView control ('lvwFiles') not found!"; throw "ListView control ('lvwFiles') could not be found." }
        Write-Host "Finished loading initial items."

        # --- Event Handlers ---
        Write-Host "Assigning event handlers."
        $controls.btnBrowse.Add_Click({
            Write-Host "Browse button clicked."
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $dialog.Description = if ($IsBackup) { "Select location to save backup" } else { "Select backup folder to restore from" }
            # Try to start browsing from the current location
            if(Test-Path $controls.txtSaveLoc.Text){ $dialog.SelectedPath = $controls.txtSaveLoc.Text } else { $dialog.SelectedPath = $script:DefaultPath }
            $dialog.ShowNewFolderButton = $IsBackup # Allow creating new folder only for backup destination

            # Use hidden owner form for better dialog behavior
            $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized'; Opacity = 0 }; $owner.Show()
            $result = $dialog.ShowDialog($owner); $owner.Close(); $owner.Dispose()

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $selectedPath = $dialog.SelectedPath; Write-Host "Folder selected: $selectedPath"
                $controls.txtSaveLoc.Text = $selectedPath; & $script:UpdateFreeSpaceLabel -ControlsParam $controls

                if (-not $IsBackup) { # Restore mode: Update list view from selected backup
                    Write-Host "Restore Mode: Loading items from selected backup folder: $selectedPath"
                    $logFilePath = Join-Path $selectedPath "FileList_Backup.csv"
                    $itemsList = [System.Collections.Generic.List[PSCustomObject]]::new() # New list for the selected backup
                    if (Test-Path -Path $logFilePath) {
                         Write-Host "Log file found. Populating ListView from log."
                         try {
                             Import-Csv $logFilePath | Where-Object { $_.BackupRelativePath -notmatch '^(ERROR_|SKIPPED_)' } | Group-Object { ($_.BackupRelativePath -split '[\\/]', 2)[0] } | ForEach-Object {
                                 $topLevelItemName = $_.Name
                                 $sourcePathInBackup = Join-Path $selectedPath $topLevelItemName
                                 $itemType = "File"; if (Test-Path $sourcePathInBackup -PathType Container -EA SilentlyContinue) { $itemType = "Folder" }
                                 $itemsList.Add([PSCustomObject]@{
                                     Name = $topLevelItemName
                                     Type = $itemType
                                     Path = $sourcePathInBackup
                                     IsSelected = $true
                                     LocalDestinationPath = ($_.Group | Select-Object -First 1).OriginalFullPath
                                 })
                             }
                             $controls.lblStatus.Content = "Ready to restore from: $selectedPath"
                             Write-Host "ListView updated with $($itemsList.Count) items from log."
                         } catch {
                             Write-Warning "Error reading backup log '$logFilePath': $($_.Exception.Message). Cannot populate list."
                             $controls.lblStatus.Content = "Error reading backup log in selected folder."; [System.Windows.MessageBox]::Show("Could not read the backup log file (FileList_Backup.csv) in the selected folder.", "Log Read Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                         }
                    } else {
                         Write-Warning "Selected folder is not a valid backup (missing FileList_Backup.csv)."
                         $controls.lblStatus.Content = "Selected folder is not a valid backup."; [System.Windows.MessageBox]::Show("Selected folder is not a valid backup (missing FileList_Backup.csv).", "Invalid Backup Folder", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    }
                    # Update the ListView with items from the newly selected backup folder
                    $controls.lvwFiles.ItemsSource = $itemsList;
                    # *** FIX: Update required space for the newly selected backup in Restore mode ***
                    & $script:UpdateRequiredSpaceLabel -ControlsParam $controls
                } else { # Backup mode: Just update status and required space
                    $controls.lblStatus.Content = "Backup location set to: $selectedPath"
                    & $script:UpdateRequiredSpaceLabel -ControlsParam $controls
                }
            } else { Write-Host "Folder selection cancelled."}
        }) # End btnBrowse Click

        $controls.btnAddFile.Add_Click({
            if (-not $IsBackup) { return } # Only works in Backup mode
            Write-Host "Add File button clicked."
            $dialog = New-Object System.Windows.Forms.OpenFileDialog; $dialog.Title = "Select File(s) to Add"; $dialog.Multiselect = $true
            $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized'; Opacity = 0 }; $owner.Show()
            $result = $dialog.ShowDialog($owner); $owner.Close(); $owner.Dispose()

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                Write-Host "$($dialog.FileNames.Count) file(s) selected."
                # Get current items, converting from potential ObservableCollection/etc. to a List
                $currentItems = [System.Collections.Generic.List[PSCustomObject]]::new()
                if ($controls.lvwFiles.ItemsSource -ne $null) { foreach ($item in $controls.lvwFiles.ItemsSource) { $currentItems.Add($item) } }
                $existingPaths = $currentItems.Path # Get list of paths already in the view

                $addedCount = 0
                foreach ($file in $dialog.FileNames) {
                    # Check if file path is already in the list (case-insensitive)
                    if (-not ($existingPaths -contains $file)) {
                        Write-Host "Adding file: $file"
                        $fileInfo = Get-Item -LiteralPath $file
                        # Create a new object matching the structure from Get-BackupPaths
                        $newItem = [PSCustomObject]@{
                            Name = $fileInfo.Name
                            Path = $fileInfo.FullName
                            Type = "File"
                            IsSelected = $true
                            LocalDestinationPath = $fileInfo.FullName # For files, local destination is the same
                        }
                        $currentItems.Add($newItem)
                        $addedCount++
                    } else { Write-Host "Skipping duplicate file: $file"}
                }
                if ($addedCount -gt 0) {
                    $controls.lvwFiles.ItemsSource = $currentItems # Update the ItemsSource with the modified list
                    Write-Host "Added $addedCount new file(s).";
                    & $script:UpdateRequiredSpaceLabel -ControlsParam $controls # Recalculate required space
                }
            } else { Write-Host "File selection cancelled."}
        }) # End btnAddFile Click

        $controls.btnAddFolder.Add_Click({
            if (-not $IsBackup) { return } # Only works in Backup mode
            Write-Host "Add Folder button clicked."
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog; $dialog.Description = "Select Folder to Add"; $dialog.ShowNewFolderButton = $false
            $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized'; Opacity = 0 }; $owner.Show()
            $result = $dialog.ShowDialog($owner); $owner.Close(); $owner.Dispose()

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                 $selectedPath = $dialog.SelectedPath; Write-Host "Folder selected to add: $selectedPath"
                 $currentItems = [System.Collections.Generic.List[PSCustomObject]]::new()
                 if ($controls.lvwFiles.ItemsSource -ne $null) { foreach ($item in $controls.lvwFiles.ItemsSource) { $currentItems.Add($item) } }
                 $existingPaths = $currentItems.Path

                 if (-not ($existingPaths -contains $selectedPath)) {
                    Write-Host "Adding folder: $selectedPath"
                    $folderInfo = Get-Item -LiteralPath $selectedPath
                    $newItem = [PSCustomObject]@{
                        Name = $folderInfo.Name
                        Path = $folderInfo.FullName
                        Type = "Folder"
                        IsSelected = $true
                        LocalDestinationPath = $folderInfo.FullName # For user-added folders, assume destination matches source
                    }
                    $currentItems.Add($newItem)
                    $controls.lvwFiles.ItemsSource = $currentItems; Write-Host "Updated ListView with new folder.";
                    & $script:UpdateRequiredSpaceLabel -ControlsParam $controls
                 } else { Write-Host "Skipping duplicate folder: $selectedPath" }
            } else { Write-Host "Folder selection cancelled."}
        }) # End btnAddFolder Click

        $controls.btnAddBAUPaths.Add_Click({
            if (-not $IsBackup) { Write-Warning "Add User Folders button disabled in Restore mode."; return }
            Write-Host "Add User Folders button clicked."
            $bauPaths = $null;
            try { $bauPaths = Get-UserPaths } catch { Write-Error "Error calling Get-UserPaths: $($_.Exception.Message)"; [System.Windows.MessageBox]::Show("Error retrieving user folders: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error); return }

            if ($bauPaths -ne $null -and $bauPaths.Count -gt 0) {
                $currentItems = [System.Collections.Generic.List[PSCustomObject]]::new()
                if ($controls.lvwFiles.ItemsSource -ne $null) { foreach ($item in $controls.lvwFiles.ItemsSource) { $currentItems.Add($item) } }
                $existingPaths = $currentItems.Path
                $addedCount = 0
                foreach ($bauItem in $bauPaths) {
                    # Get-UserPaths already checks existence, but double-check Path property
                    if ($bauItem.PSObject.Properties['Path'] -ne $null -and -not ($existingPaths -contains $bauItem.Path)) {
                        Write-Host "Adding User path: $($bauItem.Path)"
                        # Add the object directly as Get-UserPaths now includes all needed properties
                        $currentItems.Add($bauItem)
                        $addedCount++
                    } else { Write-Host "Skipping duplicate or invalid User path: $($bauItem.Path)" }
                }
                if ($addedCount -gt 0) {
                    $controls.lvwFiles.ItemsSource = $currentItems; Write-Host "Added $addedCount new User path(s).";
                    & $script:UpdateRequiredSpaceLabel -ControlsParam $controls
                } else {
                    Write-Host "No new User paths added.";
                    [System.Windows.MessageBox]::Show("No new user folders added (they may already be in the list or don't exist).", "Info", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                }
            } else { Write-Host "Get-UserPaths returned no valid paths."; [System.Windows.MessageBox]::Show("Could not find standard user folders.", "Info", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) }
        }) # End btnAddBAUPaths Click

        $controls.btnRemove.Add_Click({
            Write-Host "Remove Selected button clicked."
            # Get the items currently selected in the ListView UI
            $selectedObjects = @($controls.lvwFiles.SelectedItems)
            if ($selectedObjects.Count -gt 0) {
                Write-Host "Removing $($selectedObjects.Count) selected item(s)."
                # Get the current full list
                $currentItems = [System.Collections.Generic.List[PSCustomObject]]::new()
                if ($controls.lvwFiles.ItemsSource -ne $null) { foreach ($item in $controls.lvwFiles.ItemsSource) { $currentItems.Add($item) } }

                # Create a hash set of paths to remove for efficient lookup
                $pathsToRemove = [System.Collections.Generic.HashSet[string]]::new(
                    [string[]]($selectedObjects | Select-Object -ExpandProperty Path),
                    [System.StringComparer]::OrdinalIgnoreCase # Case-insensitive path comparison
                )

                # Filter the list, keeping only items whose paths are NOT in the removal set
                $itemsToKeep = $currentItems | Where-Object { -not $pathsToRemove.Contains($_.Path) }

                # Update the ListView's source
                $controls.lvwFiles.ItemsSource = $itemsToKeep;
                Write-Host "Kept $($itemsToKeep.Count) items.";
                # *** FIX: Recalculate space regardless of mode after removal ***
                & $script:UpdateRequiredSpaceLabel -ControlsParam $controls
            } else { Write-Host "No items selected to remove."}
        }) # End btnRemove Click

        # --- Event handlers for CheckBox clicks and Selection Changes to update Required Space ---
        # Create a script block for the update action
        # *** FIX: Remove IsBackup check - calculate space on change in both modes ***
        $updateSpaceDelegateAction = [action]{
            # Write-Host "ListView check/selection changed, updating required space." # Can be noisy
            & $script:UpdateRequiredSpaceLabel -ControlsParam $controls # Calculate space regardless of mode
        }
        # Use InvokeAsync with lower priority to avoid blocking UI thread during updates
        $updateSpaceDelegate = { param($sender, $e) $sender.Dispatcher.InvokeAsync($updateSpaceDelegateAction, [System.Windows.Threading.DispatcherPriority]::Background) | Out-Null }

        # Trigger update when selection changes (less critical, but good for consistency)
        $controls.lvwFiles.Add_SelectionChanged($updateSpaceDelegate)

        # Trigger update when a CheckBox *within* the ListView is clicked
        # PreviewMouseLeftButtonUp ensures we catch the click *before* potential selection change events
        $controls.lvwFiles.Add_PreviewMouseLeftButtonUp({
            param($sender, $e)
            # Check if the click originated from a CheckBox element
            if ($e.OriginalSource -is [System.Windows.Controls.CheckBox]) {
                # Use Dispatcher to invoke the update asynchronously
                $sender.Dispatcher.InvokeAsync($updateSpaceDelegateAction, [System.Windows.Threading.DispatcherPriority]::Background) | Out-Null
            }
        })
        Write-Host "Assigned event handlers."

        # --- Start Button Logic ---
        $controls.btnStart.Add_Click({
            $modeStringLocal = if ($IsBackup) { 'Backup' } else { 'Restore' }
            Write-Host "Start button clicked. Mode: $modeStringLocal"

            $location = $controls.txtSaveLoc.Text
            Write-Host "Selected location: $location"

            # --- Input Validation ---
            if ([string]::IsNullOrEmpty($location)) {
                [System.Windows.MessageBox]::Show("Please select a valid target directory first.", "Location Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning); return
            }
            if ($IsBackup -and -not (Test-Path $location -PathType Container)) {
                [System.Windows.MessageBox]::Show("The selected backup destination directory does not exist or is not accessible: `n$location", "Backup Location Invalid", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning); return
            }
            # For restore, check for the essential log file in the selected directory
            if (-not $IsBackup -and -not (Test-Path (Join-Path $location "FileList_Backup.csv") -PathType Leaf)) {
                [System.Windows.MessageBox]::Show("The selected restore source is not a valid backup folder (missing FileList_Backup.csv): `n$location", "Restore Source Invalid", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning); return
            }

            # *** FIX: Recalculate required space just before starting operation (regardless of mode) ***
            Write-Host "Recalculating required space before starting operation...";
            & $script:UpdateRequiredSpaceLabel -ControlsParam $controls
            # Optional: Add a check against free space here if desired (more relevant for Backup)

            # Get the list of items that are currently CHECKED in the UI
            $itemsFromUIList = [System.Collections.Generic.List[PSCustomObject]]::new()
            if ($controls.lvwFiles.ItemsSource -ne $null) {
                # Filter the ItemsSource based on the IsSelected property bound to the CheckBox
                try { # Handle potential errors if ItemsSource is modified while iterating
                    @($controls.lvwFiles.ItemsSource) | Where-Object { $_.PSObject.Properties['IsSelected'] -ne $null -and $_.IsSelected } | ForEach-Object { $itemsFromUIList.Add($_) }
                } catch { Write-Warning "Error reading selected items from ListView: $($_.Exception.Message)" }
            }


            if ($itemsFromUIList.Count -eq 0) {
                [System.Windows.MessageBox]::Show("No items are selected (checked) to $modeStringLocal.", "No Items Selected", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning); return
            }

            # Get checkbox state - used directly for Backup, ignored for Restore decision below
            $doNetwork = $controls.chkNetwork.IsChecked
            $doPrinters = $controls.chkPrinters.IsChecked

            # --- Disable UI and Start Operation ---
            Write-Host "Disabling UI controls and setting wait cursor."
            $controls.Values | Where-Object { $_ -is [System.Windows.Controls.Control] } | ForEach-Object { $_.IsEnabled = $false }
            $window.Cursor = [System.Windows.Input.Cursors]::Wait
            $controls.prgProgress.IsIndeterminate = $true # Start as indeterminate
            $controls.prgProgress.Value = 0
            $controls.txtProgress.Text = "Initializing..."
            $controls.lblStatus.Content = "Starting $modeStringLocal..."

            # Define the progress action to update the UI from the background operation
            $uiProgressAction = {
                param($status, $percent, $details)
                # Ensure UI updates happen on the UI thread using Dispatcher
                if ($window.Dispatcher) {
                    $window.Dispatcher.InvokeAsync(
                        [action]{
                            $controls.lblStatus.Content = $status
                            $controls.txtProgress.Text = $details
                            # Handle progress bar state: Indeterminate (-2), Error (-1), Normal (0-100)
                            if ($percent -eq -2) { # Explicit indeterminate request
                                $controls.prgProgress.IsIndeterminate = $true
                                $controls.prgProgress.Value = 0
                            } elseif ($percent -lt 0) { # Error state
                                $controls.prgProgress.IsIndeterminate = $false
                                $controls.prgProgress.Value = 0
                                # Optionally change progress bar color for error
                                # $controls.prgProgress.Foreground = [System.Windows.Media.Brushes]::Red
                            } else { # Normal progress
                                $controls.prgProgress.IsIndeterminate = $false
                                $controls.prgProgress.Value = $percent
                                # Reset color if it was changed for error
                                # $controls.prgProgress.ClearValue([System.Windows.Controls.ProgressBar]::ForegroundProperty)
                            }
                        },
                        [System.Windows.Threading.DispatcherPriority]::Background # Use Background priority for responsiveness
                    ) | Out-Null
                }
            } # End uiProgressAction

            $success = $false
            try { # INNER try for the core Backup/Restore operation
                if ($IsBackup) {
                    # Construct unique backup folder name
                    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                    $username = $env:USERNAME -replace '[^a-zA-Z0-9_.-]', '_' # Sanitize username for path
                    $backupRootPath = Join-Path $location "Backup_${username}_$timestamp"

                    # Call the backup function - uses checkbox state for $doNetwork/$doPrinters
                    $success = Invoke-BackupOperation -BackupRootPath $backupRootPath `
                                                     -ItemsToBackup $itemsFromUIList `
                                                     -BackupNetworkDrives $doNetwork `
                                                     -BackupPrinters $doPrinters `
                                                     -ProgressAction $uiProgressAction `
                                                     -SettingsSourceComputer $env:COMPUTERNAME # Source is local machine
                } else { # Restore
                    $backupRootPath = $location # Location is the selected backup folder

                    # Determine whether to restore drives/printers based on file existence in the backup folder
                    $restoreDrivesFromFile = Test-Path (Join-Path $backupRootPath "Drives.csv") -PathType Leaf
                    $restorePrintersFromFile = Test-Path (Join-Path $backupRootPath "Printers.txt") -PathType Leaf
                    Write-Host "Starting Restore. Restore Drives (File Exists): $restoreDrivesFromFile, Restore Printers (File Exists): $restorePrintersFromFile"

                    # NOTE: Invoke-RestoreOperation currently ignores the $itemsFromUIList because it uses RestoreAllFromLog=$true.
                    # If selective restore based on UI checks was needed, Invoke-RestoreOperation would need modification
                    # to accept the $itemsFromUIList and process only those entries from the log.
                    # For now, it restores everything valid found in the log file.
                    $restoreParams = @{
                        BackupRootPath        = $backupRootPath
                        RestoreAllFromLog     = $true # Switch parameter needs explicit $true
                        RestoreNetworkDrives  = $restoreDrivesFromFile
                        RestorePrinters       = $restorePrintersFromFile
                        ProgressAction        = $uiProgressAction
                    }
                    # Call the restore function using the splatted parameters
                    $success = Invoke-RestoreOperation @restoreParams
                }

                # --- Handle Operation Result ---
                if ($success) {
                    Write-Host "Operation completed successfully (UI)." -ForegroundColor Green
                    $controls.lblStatus.Content = "Operation completed successfully."
                    $controls.txtProgress.Text = "Finished."
                    $controls.prgProgress.Value = 100
                    $controls.prgProgress.IsIndeterminate = $false
                    [System.Windows.MessageBox]::Show("The $modeStringLocal operation completed successfully!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                    # Start background tasks only after successful Restore
                    if (-not $IsBackup) { Start-BackgroundUpdateJob }
                } else {
                    Write-Error "Operation failed (UI)."
                    $controls.lblStatus.Content = "Operation Failed!"
                    # Progress action should have set error details in txtProgress
                    $controls.prgProgress.Value = 0
                    $controls.prgProgress.IsIndeterminate = $false
                    [System.Windows.MessageBox]::Show("The $modeStringLocal operation failed. Check console output or log files for details.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                }
            } catch { # INNER catch for unexpected errors during the operation call or UI updates
                # This is where the "Unrecognized term" error was likely caught
                $errorMessage = "Operation Failed (UI Catch): $($_.Exception.Message)"
                Write-Error $errorMessage
                $controls.lblStatus.Content = "Operation Failed!"
                $controls.txtProgress.Text = $errorMessage
                $controls.prgProgress.Value = 0
                $controls.prgProgress.IsIndeterminate = $false
                [System.Windows.MessageBox]::Show($errorMessage, "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            } finally { # INNER finally to re-enable UI
                Write-Host "Operation finished (inner finally block). Re-enabling UI controls."
                 # Re-enable controls safely using Dispatcher
                 if ($window.Dispatcher) {
                     $window.Dispatcher.InvokeAsync({
                         $controls.Values | Where-Object { $_ -is [System.Windows.Controls.Control] } | ForEach-Object {
                             # Special handling for buttons only enabled in Backup mode
                             if ($_.Name -in @('btnAddFile', 'btnAddFolder', 'btnAddBAUPaths')) {
                                 $_.IsEnabled = $IsBackup
                             } else {
                                 $_.IsEnabled = $true
                             }
                         }
                         $window.Cursor = [System.Windows.Input.Cursors]::Arrow # Reset cursor
                     }) | Out-Null
                 } else { # Fallback if dispatcher isn't available (shouldn't happen)
                     $controls.Values | Where-Object { $_ -is [System.Windows.Controls.Control] } | ForEach-Object { $_.IsEnabled = $true }
                     $window.Cursor = [System.Windows.Input.Cursors]::Arrow
                 }
                 Write-Host "UI controls re-enabled and cursor reset."
            } # End INNER finally
        }) # End btnStart.Add_Click

        Write-Host "Showing main window."
        # Show the window; script execution pauses here until the window is closed.
        $window.ShowDialog() | Out-Null
        Write-Host "Main window closed."

    } catch { # OUTER catch for errors during XAML loading or initial setup
        $errorMessage = "Failed to load or initialize main window: $($_.Exception.Message)"
        Write-Error $errorMessage; Write-Host $errorMessage -ForegroundColor Red
        try { [System.Windows.MessageBox]::Show($errorMessage, "Critical UI Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) } catch {}
    } finally { # OUTER finally for cleanup
         Write-Host "Exiting Show-MainWindow function."
         # Clean up script-scoped helper functions if they exist
         if (Test-Path variable:script:UpdateFreeSpaceLabel) { Remove-Variable -Name UpdateFreeSpaceLabel -Scope Script -ErrorAction SilentlyContinue }
         if (Test-Path variable:script:UpdateRequiredSpaceLabel) { Remove-Variable -Name UpdateRequiredSpaceLabel -Scope Script -ErrorAction SilentlyContinue }
    } # End OUTER finally
} # End Show-MainWindow function


#endregion Functions

# --- Main Execution ---
Write-Host "--- Script Starting ---"
# Clear any previous background job variable
Clear-Variable -Name updateJob -Scope Script -ErrorAction SilentlyContinue

# Ensure Default Path Exists early on
if (-not (Test-Path $script:DefaultPath)) {
    Write-Host "Default path '$($script:DefaultPath)' not found. Attempting to create."
    try { New-Item -Path $script:DefaultPath -ItemType Directory -Force -ErrorAction Stop | Out-Null; Write-Host "Default path created." }
    catch { Write-Warning "Could not create default path: $($script:DefaultPath). Express and Remote modes might fail if they rely on it." }
}

# Define Console Progress Action for non-GUI modes
$consoleProgressAction = {
    param($status, $percent, $details)
     $percentString = if ($percent -lt 0) { "ERR" } elseif ($percent -eq -2) { "..." } else { "$percent%" }
     Write-Host "[$status - $percentString]: $details"
     # Update console progress bar if running in console
     if ($Host.Name -eq 'ConsoleHost') {
         try {
             if ($percent -ge 0) { Write-Progress -Activity $status -Status "$percent% Complete" -CurrentOperation $details -PercentComplete $percent -Id 1 }
             elseif ($percent -eq -2) { Write-Progress -Activity $status -Status "Processing..." -CurrentOperation $details -PercentComplete 0 -Id 1 } # Show 0% for indeterminate
             else { Write-Progress -Activity $status -Status "Error!" -CurrentOperation $details -PercentComplete 0 -Id 1 } # Show 0% for error
         } catch {
             # Write-Progress can sometimes fail (e.g., in certain remote sessions or hosts)
             Write-Warning "Failed to update console progress bar: $($_.Exception.Message)"
         }
     }
}

# --- Mode Selection and Elevation Handling ---
$script:selectedMode = $null
if ($ElevatedRemoteRun) {
    # If script was re-launched with elevation for Remote mode
    Write-Host "Script re-launched with elevated privileges for Remote mode." -ForegroundColor Yellow
    $script:selectedMode = 'Remote'
} else {
    # Show the initial mode selection dialog
    Write-Host "Calling Show-ModeDialog to determine operation mode."
    $script:selectedMode = Show-ModeDialog
}

# --- Conditional Elevation Check for Remote Mode ---
if ($script:selectedMode -eq 'Remote' -and -not (Test-IsAdmin)) {
    Write-Warning "Remote mode requires administrator privileges."
    Write-Host "Attempting to restart the script with elevation..."
    try {
        # Get path to powershell.exe
        $powershellPath = Get-Command powershell.exe | Select-Object -ExpandProperty Source
        # Get path to the current script
        $scriptPath = $MyInvocation.MyCommand.Path
        # Define arguments for the elevated process
        $processArgs = "-NoProfile -ExecutionPolicy Bypass -File ""$scriptPath"" -ElevatedRemoteRun"
        # Start the process elevated
        Start-Process -FilePath $powershellPath -ArgumentList $processArgs -Verb RunAs -ErrorAction Stop
        Write-Host "Elevated process started. Exiting current non-elevated instance."
        Exit # Exit the current non-elevated script
    } catch {
        $errorMsg = "Failed to restart script with administrator privileges. Please run the script manually as an administrator to use Remote mode. Error: $($_.Exception.Message)"
        Write-Error $errorMsg
        # Show message box as fallback if console is closing
        try { [System.Windows.MessageBox]::Show($errorMsg, "Elevation Failed", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) } catch {}
        Exit 1 # Exit with an error code
    }
}

# --- Main Logic Execution Based on Selected Mode ---
try {
    switch ($script:selectedMode) {
        'Backup' {
            Write-Host "Mode selected: Backup. Showing main window."
            # Call the GUI function for Backup mode
            Show-MainWindow -IsBackup $true
        }
        'Restore' {
            Write-Host "Mode selected: Restore. Showing main window."
            # Call the GUI function for Restore mode
            Show-MainWindow -IsBackup $false
        }
        'Express' {
            Write-Host "Mode selected: Express. Executing Express logic..." -ForegroundColor Cyan
            Write-Host "Checking for recent backup in '$($script:DefaultPath)'..."
            $todayDate = (Get-Date).Date
            # Find the latest backup folder based on LastWriteTime (more reliable than name)
            $latestBackup = Get-ChildItem -Path $script:DefaultPath -Directory -Filter "Backup_*" |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1

            $restoreCandidatePath = $null
            # Check if a backup exists AND if it was created today
            if ($latestBackup -and $latestBackup.CreationTime.Date -eq $todayDate) {
                Write-Host "Recent backup found from today: $($latestBackup.FullName)" -ForegroundColor Green
                $restoreCandidatePath = $latestBackup.FullName
            } else {
                if ($latestBackup) { Write-Host "Latest backup found ($($latestBackup.Name)) is not from today ($($latestBackup.CreationTime.ToString('yyyy-MM-dd')))." }
                else { Write-Host "No existing backups found in '$($script:DefaultPath)'." }
            }

            # --- Express Restore ---
            if ($restoreCandidatePath) {
                Write-Host "Starting Express Restore from '$restoreCandidatePath'..." -ForegroundColor Yellow
                # Check if optional components exist in the backup for the restore operation
                $restoreDrives = Test-Path (Join-Path $restoreCandidatePath "Drives.csv") -PathType Leaf
                $restorePrinters = Test-Path (Join-Path $restoreCandidatePath "Printers.txt") -PathType Leaf
                Write-Host "Restore Options - Drives File Exists: $restoreDrives, Printers File Exists: $restorePrinters"

				$restoreExpressParams = @{
						BackupRootPath        = $restoreCandidatePath
						RestoreAllFromLog     = $true # Switch parameter needs explicit $true
						RestoreNetworkDrives  = $restoreDrives
						RestorePrinters       = $restorePrinters
						ProgressAction        = $consoleProgressAction
					}
					# Call the restore function using the splatted parameters
					$success = Invoke-RestoreOperation @restoreExpressParams

                if ($success) {
                    Write-Host "Express Restore completed successfully." -ForegroundColor Green
                    Start-BackgroundUpdateJob # Run GPUpdate/CM Actions after restore
                    try { [System.Windows.MessageBox]::Show("Express Restore completed successfully from `n$restoreCandidatePath", "Express Restore Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) } catch {}
                } else {
                    Write-Error "Express Restore failed."
                    try { [System.Windows.MessageBox]::Show("Express Restore failed. Check console output for details.", "Express Restore Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) } catch {}
                }
            }
            # --- Express Backup ---
            else {
                Write-Host "No recent backup found. Starting Express Backup to '$($script:DefaultPath)'..." -ForegroundColor Yellow
                $itemsToBackupList = [System.Collections.Generic.List[PSCustomObject]]::new()
                # Gather default and user paths
                try { # FIX: Use iterative Add instead of AddRange
                    $backupPaths = Get-BackupPaths -UserProfilePath $env:USERPROFILE
                    if ($backupPaths) { foreach ($item in $backupPaths) { $itemsToBackupList.Add($item) } }
                } catch { Write-Warning "Error getting default backup paths: $($_.Exception.Message)"}


                # Ensure all items are marked as selected (Get-* functions should already do this, but belt-and-suspenders)
                # Also ensure LocalDestinationPath exists (should be handled by Get-* functions now)
                $itemsToBackupList | ForEach-Object {
                    # FIX: Use Add-Member -Force to ensure IsSelected exists and is true
                    $_ | Add-Member -MemberType NoteProperty -Name 'IsSelected' -Value $true -Force

                    # Verify LocalDestinationPath (Get-* functions should provide this)
                    if (-not $_.PSObject.Properties.Name -contains 'LocalDestinationPath' -or [string]::IsNullOrEmpty($_.LocalDestinationPath)) {
                         Write-Warning "Item '$($_.Name)' missing LocalDestinationPath after Get-* call. Using source Path '$($_.Path)' as fallback."
                         $_ | Add-Member -MemberType NoteProperty -Name 'LocalDestinationPath' -Value $_.Path -Force
                    }
                }

                # Filter out items where the source path doesn't actually exist
                $validItemsToBackup = $itemsToBackupList | Where-Object { Test-Path -LiteralPath $_.Path }
                $skippedCount = $itemsToBackupList.Count - $validItemsToBackup.Count
                if ($skippedCount -gt 0) {
                    Write-Warning "$skippedCount item(s) skipped because their source path does not exist."
                }

                if ($validItemsToBackup.Count -eq 0) {
                    Write-Error "Express Backup: No valid items found to back up after checking paths."
                    try { [System.Windows.MessageBox]::Show("Express Backup failed: No valid files or folders found to back up.", "Express Backup Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) } catch {}
                } else {
                    $backupDrives = $true; $backupPrinters = $true # Include drives/printers in Express Backup
                    Write-Host "Backup Options - Drives: $backupDrives, Printers: $backupPrinters"
                    Write-Host "Items to back up:"
                    $validItemsToBackup | Format-Table Name, Type, Path, LocalDestinationPath -AutoSize

                    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                    $username = $env:USERNAME -replace '[^a-zA-Z0-9_.-]', '_'
                    $backupRootPath = Join-Path $script:DefaultPath "Backup_${username}_$timestamp"

                    $success = Invoke-BackupOperation -BackupRootPath $backupRootPath `
                                                     -ItemsToBackup $validItemsToBackup `
                                                     -BackupNetworkDrives $backupDrives `
                                                     -BackupPrinters $backupPrinters `
                                                     -ProgressAction $consoleProgressAction `
                                                     -SettingsSourceComputer $env:COMPUTERNAME
                    if ($success) {
                        Write-Host "Express Backup completed successfully." -ForegroundColor Green
                        try { [System.Windows.MessageBox]::Show("Express Backup completed successfully to `n$backupRootPath", "Express Backup Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) } catch {}
                    } else {
                        Write-Error "Express Backup failed."
                        try { [System.Windows.MessageBox]::Show("Express Backup failed. Check console output for details.", "Express Backup Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) } catch {}
                    }
                } # End else ($validItemsToBackup.Count -gt 0)
            } # End else (Express Backup)
        } # End Express Case
        'Remote' {
            # This block requires elevation (checked earlier)
            Write-Host "Mode selected: Remote. Executing Remote logic (Elevated)." -ForegroundColor Cyan
            $targetDevice = $null
            $remoteUsername = $null
            $mappedDrivePath = "$($script:RemoteMappedDriveLetter):"
            $localBackupPath = $null
            $remoteSuccess = $false # Flag to track overall success

            try { # Main try block for the entire remote operation
                # --- Get Target Device ---
                while ([string]::IsNullOrWhiteSpace($targetDevice)) {
                    $targetDevice = Read-Host "Enter the IP address or Hostname of the OLD (source) device"
                    if ([string]::IsNullOrWhiteSpace($targetDevice)) { Write-Warning "Target device cannot be empty." }
                    # Basic ping test for quick validation
                    elseif (-not (Test-Connection -ComputerName $targetDevice -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
                        Write-Warning "Cannot ping target device '$targetDevice'. Please check name/IP and network connectivity."
                        $targetDevice = $null # Force re-entry
                    }
                }
                Write-Host "Target device: $targetDevice"

                # --- Map Network Drive ---
                $uncPath = "\\$targetDevice\c$"
                Write-Host "Attempting to map network drive '$uncPath' to '$mappedDrivePath'..."
                if (Test-Path $mappedDrivePath) {
                    Write-Warning "Drive $mappedDrivePath already exists. Attempting to remove..."
                    try { Remove-PSDrive -Name $script:RemoteMappedDriveLetter -Force -ErrorAction Stop }
                    catch { throw "Failed to remove existing mapped drive '$mappedDrivePath'. Please remove it manually and retry. Error: $($_.Exception.Message)" }
                    Write-Host "Removed existing drive $mappedDrivePath."
                }
                try {
                    # Map the drive. Credentials might be required if not domain-joined or different creds needed.
                    # For simplicity, this assumes current elevated user has access. Add -Credential if needed.
                    New-PSDrive -Name $script:RemoteMappedDriveLetter -PSProvider FileSystem -Root $uncPath -ErrorAction Stop | Out-Null
                    Write-Host "Successfully mapped '$uncPath' to '$mappedDrivePath'." -ForegroundColor Green
                } catch {
                    throw "Failed to map network drive '$uncPath'. Verify connectivity, firewall settings (File and Printer Sharing), administrative share (c$) access, and permissions on '$targetDevice'. Error: $($_.Exception.Message)"
                }

                # --- Get Remote Username ---
                Write-Host "Attempting to detect username on '$targetDevice'..."
                try {
                    # Try CIM first (more modern)
                    $remoteCompSys = Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $targetDevice -ErrorAction SilentlyContinue
                    if ($remoteCompSys -and $remoteCompSys.UserName) {
                        $remoteUsername = ($remoteCompSys.UserName -split '\\')[-1] # Get username part
                        Write-Host "Auto-detected remote user (via CIM): $remoteUsername" -ForegroundColor Green
                    } else {
                        # Fallback to WMI if CIM fails or returns no user
                        Write-Warning "CIM query failed or returned no username. Trying WMI..."
                        $remoteCompSysWMI = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $targetDevice -ErrorAction Stop
                        if ($remoteCompSysWMI -and $remoteCompSysWMI.UserName) {
                             $remoteUsername = ($remoteCompSysWMI.UserName -split '\\')[-1]
                             Write-Host "Auto-detected remote user (via WMI): $remoteUsername" -ForegroundColor Green
                        } else { throw "Could not retrieve username from remote machine via CIM or WMI." }
                    }
                } catch {
                    # Manual prompt if auto-detection fails
                    Write-Warning "Failed to auto-detect remote user: $($_.Exception.Message)"
                    while ([string]::IsNullOrWhiteSpace($remoteUsername)) {
                        $remoteUsername = Read-Host "Could not auto-detect. Please enter the Windows username logged into '$targetDevice'"
                        if ([string]::IsNullOrWhiteSpace($remoteUsername)) { Write-Warning "Remote username cannot be empty." }
                    }
                    Write-Host "Using manually entered remote username: $remoteUsername"
                }

                # --- Get Items from Remote Profile ---
                $remoteUserProfilePath = Join-Path $mappedDrivePath "Users\$remoteUsername"
                Write-Host "Checking remote user profile path: '$remoteUserProfilePath'"
                if (-not (Test-Path -LiteralPath $remoteUserProfilePath -PathType Container)) {
                    throw "Remote user profile path '$remoteUserProfilePath' not found on mapped drive. Verify the username ('$remoteUsername') and profile existence on '$targetDevice'."
                }

                Write-Host "Gathering items to copy from remote profile '$remoteUserProfilePath'..."
                # Use Get-BackupPaths, providing the path to the *remote* profile
                $itemsToCopyList = Get-BackupPaths -UserProfilePath $remoteUserProfilePath
                # Optionally add standard user folders from remote drive too
                # $itemsToCopyList.AddRange( (Get-UserPaths -UserProfilePath $remoteUserProfilePath) ) # Needs Get-UserPaths modification to accept path

                if (-not $itemsToCopyList -or $itemsToCopyList.Count -eq 0) {
                    # This isn't necessarily fatal, maybe only drives/printers are needed
                    Write-Warning "No specific backup items (Signatures, Bookmarks, etc.) found in the standard locations for user '$remoteUsername' on '$targetDevice'."
                } else {
                    Write-Host "Found $($itemsToCopyList.Count) items/folders to potentially copy from remote."
                    $itemsToCopyList | Format-Table Name, Type, Path, LocalDestinationPath -AutoSize
                }

                # --- Perform Remote Backup to Local Folder ---
                $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                $safeTargetName = $targetDevice -replace '[^a-zA-Z0-9_.-]', '_'
                $safeRemoteUser = $remoteUsername -replace '[^a-zA-Z0-9_.-]', '_'
                $localBackupPath = Join-Path $script:DefaultPath "RemoteBackup_${safeRemoteUser}_from_${safeTargetName}_$timestamp"
                Write-Host "Creating local backup folder: '$localBackupPath'"
                New-Item -Path $localBackupPath -ItemType Directory -Force -ErrorAction Stop | Out-Null

                Write-Host "Starting backup process from remote source '$targetDevice' to local folder '$localBackupPath'..."
                # Call Invoke-BackupOperation:
                # - ItemsToBackup uses paths on the mapped drive (e.g., Z:\Users\...)
                # - SettingsSourceComputer is the remote device name for getting drives/printers
                $backupSuccess = Invoke-BackupOperation -BackupRootPath $localBackupPath `
                                                      -ItemsToBackup $itemsToCopyList `
                                                      -BackupNetworkDrives $true `
                                                      -BackupPrinters $true `
                                                      -ProgressAction $consoleProgressAction `
                                                      -SettingsSourceComputer $targetDevice

                if (-not $backupSuccess) {
                    # Don't stop necessarily, maybe only files failed but drives/printers worked. Logged by Invoke-BackupOperation.
                    Write-Warning "Backup operation from remote source completed with errors. Check logs above. Continuing with restore attempt..."
                } else {
                    Write-Host "Remote data and settings successfully backed up locally to '$localBackupPath'." -ForegroundColor Green
                }

                # --- Perform Local Restore ---
                Write-Host "Starting local restore from '$localBackupPath'..."
                # Check for drive/printer files in the newly created local backup
                $restoreDrives = Test-Path (Join-Path $localBackupPath "Drives.csv") -PathType Leaf
                $restorePrinters = Test-Path (Join-Path $localBackupPath "Printers.txt") -PathType Leaf
                Write-Host "Local Restore Check - Drives File Exists: $restoreDrives, Printers File Exists: $restorePrinters"

                $restoreSuccess = Invoke-RestoreOperation -BackupRootPath $localBackupPath `
                                                        -RestoreAllFromLog `
                                                        -RestoreNetworkDrives $restoreDrives ` # Pass file check result
                                                        -RestorePrinters $restorePrinters `    # Pass file check result
                                                        -ProgressAction $consoleProgressAction

                if (-not $restoreSuccess) {
                    # Logged by Invoke-RestoreOperation
                    Write-Warning "Local restore operation completed with errors. Check logs above."
                } else {
                    Write-Host "Local restore completed successfully." -ForegroundColor Green
                }

                # --- Start Post-Restore Tasks ---
                Write-Host "Starting background job for local GPUpdate and ConfigMgr actions..."
                Start-BackgroundUpdateJob

                $remoteSuccess = $true # Mark overall success if we got this far (even with warnings)

            } catch { # Catch errors during the remote process
                $errorMessage = "An error occurred during the Remote operation: $($_.Exception.Message)"
                Write-Error $errorMessage; Write-Host $errorMessage -ForegroundColor Red
                try { [System.Windows.MessageBox]::Show($errorMessage, "Remote Mode Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) } catch {}
                $remoteSuccess = $false # Mark as failed
            } finally {
                # --- Cleanup Mapped Drive ---
                if (Test-Path $mappedDrivePath) {
                    Write-Host "Cleaning up: Removing mapped drive '$mappedDrivePath'..."
                    try { Remove-PSDrive -Name $script:RemoteMappedDriveLetter -Force -ErrorAction Stop }
                    catch { Write-Warning "Failed to remove mapped drive '$mappedDrivePath': $($_.Exception.Message)" }
                }

                # --- Final Status Message ---
                if ($remoteSuccess) {
                    Write-Host "Remote Mode process finished. Check console for details and background job output." -ForegroundColor Green
                    try { [System.Windows.MessageBox]::Show("Remote data transfer and local restore process finished.`nBackup created at: $localBackupPath`nCheck console for details and background job output.", "Remote Mode Finished", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) } catch {}
                } else {
                    Write-Error "Remote Mode process failed or was interrupted. Please review console output."
                }
                 # Ensure console progress bar is cleared
                 if ($Host.Name -eq 'ConsoleHost') { try { Write-Progress -Activity "Remote Operation" -Completed -Id 1 -ErrorAction SilentlyContinue } catch {} }
            }
        } # End Remote Case
        'Cancel' {
            Write-Host "Operation cancelled by user."
        }
        Default {
            # This case handles unexpected results from Show-ModeDialog or if $ElevatedRemoteRun was true but mode wasn't 'Remote'
            if (-not $ElevatedRemoteRun) { # Avoid warning if it was just the elevated run finishing
                 Write-Warning "Invalid mode selected or operation cancelled: '$script:selectedMode'"
            }
        }
    } # End Switch

} catch { # Top-level catch for unexpected script errors
    $errorMessage = "An unexpected critical error occurred: $($_.Exception.Message)"
    Write-Error $errorMessage; Write-Host $errorMessage -ForegroundColor Red
    try { [System.Windows.MessageBox]::Show($errorMessage, "Fatal Script Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) } catch {}
} finally {
    # --- Wait for and Display Background Job Output ---
    # Check if the job variable exists and is a job object
    if ((Get-Variable -Name updateJob -Scope Script -ErrorAction SilentlyContinue) -ne $null -and $script:updateJob -is [System.Management.Automation.Job]) {
        Write-Host "`n--- Waiting for background update job (GPUpdate/CM Actions) to complete... ---" -ForegroundColor Yellow
        # Wait for the job to finish
        Wait-Job $script:updateJob | Out-Null
        Write-Host "--- Background Update Job Output (GPUpdate/CM Actions): ---" -ForegroundColor Yellow
        # Retrieve and display the job's output
        Receive-Job $script:updateJob
        # Clean up the job object
        Remove-Job $script:updateJob
        Write-Host "--- End of Background Update Job Output ---" -ForegroundColor Yellow
    } else {
        Write-Host "`nNo background update job was started or it was already cleaned up." -ForegroundColor Gray
    }
}

Write-Host "--- Script Execution Finished ---"
# Pause console window only if run directly in console, not ISE/VSCode, and not during elevated relaunch
if ($Host.Name -eq 'ConsoleHost' -and -not $psISE -and $env:TERM_PROGRAM -ne 'vscode' -and -not $ElevatedRemoteRun) {
    Write-Host "Press Enter to exit..." -ForegroundColor Yellow
    Read-Host
}