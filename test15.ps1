# Main.ps1 - works but the user backup button isn't working. 

##############################################################################################################################
#        Client Data Backup Tool         											#
#       Written By Jared Vosters - ACT Logistics, github.com/cattboy        		#
# Based off the script by Stephen Onions & Kevin King UserBackupRefresh_Persist 1	#
#
# This Powershell script runs locally only, has 3 prompts. EXPRESS, BACKUP, RESTORE
#
# Express - Attempts to run without initial admin rights. If selected, checks for elevation.
#           If not elevated, prompts user and attempts to relaunch script as Admin.
#           Once elevated: Prompts for OLD device name/IP.
#           Attempts to map C$ drive of OLD device.
#           Transfers specified user data/settings from OLD device to NEW (local) device.
#           Captures/recreates network drives & printers from OLD device to NEW.
#           Creates a backup copy of the transferred data in C:\LocalData on NEW device.
#
# Backup - GUI interface for local backup. Runs without admin.
#
# Restore - GUI interface for local restore from a backup folder. Runs without initial admin.
#           Background job for GPUpdate/SCCM Actions may require elevation to succeed fully.
#
###########################################################################################################################
#requires -Version 3.0
# Removed: #requires -RunAsAdministrator

# --- Add Param block at the TOP ---
param(
    # Switch parameter to indicate the script was relaunched elevated for Express mode
    [switch]$RunExpressElevated
)

# --- Load Assemblies FIRST ---
Write-Host "Attempting to load .NET Assemblies..."
try {
    Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop
    Write-Host ".NET Assemblies loaded successfully."
} catch {
    Write-Error "FATAL: Failed to load required .NET Assemblies. UI cannot be created."
    Write-Error "Error Message: $($_.Exception.Message)"
    if ($Host.Name -eq 'ConsoleHost' -and -not $psISE -and $env:TERM_PROGRAM -ne 'vscode') { Read-Host "Press Enter to exit" }
    Exit 1
}

#region Global Variables & Constants
$script:DefaultPath = "C:\LocalData"
$script:RemoteDriveLetter = "X" # Choose a drive letter unlikely to be in use
#endregion

#region Functions

# --- Helper Functions ---

function Test-IsAdmin {
    try {
        $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        Write-Warning "Error checking admin status: $($_.Exception.Message)"
        return $false
    }
}

function Show-InputBox {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Prompt,
        [string]$Title = "Input Required",
        [string]$DefaultText = ""
    )
    try {
        return [Microsoft.VisualBasic.Interaction]::InputBox($Prompt, $Title, $DefaultText)
    } catch {
        Write-Error "Failed to show InputBox: $($_.Exception.Message)"
        return $null
    }
}

function Format-Bytes {
    param([Parameter(Mandatory=$true)][long]$Bytes)
    $suffix = @("B", "KB", "MB", "GB", "TB", "PB"); $index = 0; $value = [double]$Bytes
    while ($value -ge 1024 -and $index -lt ($suffix.Length - 1)) { $value /= 1024; $index++ }
    return "{0:N2} {1}" -f $value, $suffix[$index]
}

# --- Core Logic Functions (Local Backup/Restore) ---

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

    Write-Host "--- Starting LOCAL Backup Operation to '$BackupRootPath' ---"
    $UpdateProgress = { if ($ProgressAction) { $ProgressAction.Invoke($args[0], $args[1], $args[2]) } else { Write-Host "$($args[0]) ($($args[1])%) - $($args[2])" } }

    try { # Line 48 - Start of try block
        $UpdateProgress.Invoke("Initializing Backup", 0, "Creating backup directory...")
        if (-not (Test-Path $BackupRootPath)) {
            New-Item -Path $BackupRootPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Write-Host "Backup directory created: $BackupRootPath"
        }

        $csvLogPath = Join-Path $BackupRootPath "FileList_Backup.csv"
        Write-Host "Creating log file: $csvLogPath"
        # Log format for local backup/restore
        "OriginalFullPath,BackupRelativePath" | Set-Content -Path $csvLogPath -Encoding UTF8

        if (-not $ItemsToBackup -or $ItemsToBackup.Count -eq 0) {
             throw "No items specified for backup."
        }
        Write-Host "Processing $($ItemsToBackup.Count) items for backup."

        # Estimate total files for progress
        $UpdateProgress.Invoke("Estimating size...", 0, "Calculating total files...")
        $totalFilesEstimate = 0
        $itemsToProcessForSize = $ItemsToBackup
        $itemsToProcessForSize | ForEach-Object {
            if (Test-Path -LiteralPath $_.Path -ErrorAction SilentlyContinue) {
                if ($_.Type -eq 'Folder') {
                    try { $totalFilesEstimate += (Get-ChildItem -LiteralPath $_.Path -Recurse -File -Force -ErrorAction SilentlyContinue).Count } catch {}
                } else { $totalFilesEstimate++ }
            }
        }
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
                        # For local backup, relative path is within the item's folder structure
                        $relativeFilePath = $originalFileFullPath.Substring($basePathLength).TrimStart('\')
                        $backupRelativePath = Join-Path $item.Name $relativeFilePath
                        $targetBackupPath = Join-Path $BackupRootPath $backupRelativePath
                        $targetBackupDir = [System.IO.Path]::GetDirectoryName($targetBackupPath)

                        if (-not (Test-Path $targetBackupDir)) {
                            New-Item -Path $targetBackupDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
                        }
                        Copy-Item -LiteralPath $originalFileFullPath -Destination $targetBackupPath -Force -ErrorAction Stop
                        # Log original path and path relative to backup root
                        "`"$originalFileFullPath`",`"$backupRelativePath`"" | Add-Content -Path $csvLogPath -Encoding UTF8
                        $UpdateProgress.Invoke($statusMessage, $percentComplete, "Copied: $($_.Name)")
                    }
                    Write-Host "  Finished processing folder: $($item.Name)"
                } catch { Write-Warning "Error processing folder '$($item.Name)' ($sourcePath): $($_.Exception.Message)" }
            } else { # Single File
                 Write-Host "  Item is a file. Processing..."
                 try {
                    $originalFileFullPath = $sourcePath
                    # For single files, backup relative path is just the filename
                    $backupRelativePath = $item.Name
                    $targetBackupPath = Join-Path $BackupRootPath $backupRelativePath
                    Copy-Item -LiteralPath $originalFileFullPath -Destination $targetBackupPath -Force -ErrorAction Stop
                    # Log original path and path relative to backup root
                    "`"$originalFileFullPath`",`"$backupRelativePath`"" | Add-Content -Path $csvLogPath -Encoding UTF8
                 } catch { Write-Warning "Error processing file '$($item.Name)' ($sourcePath): $($_.Exception.Message)" }
            }
        } # End foreach item

        # Backup Network Drives (Locally)
        if ($BackupNetworkDrives) {
            $currentItemIndex++
            $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 100) } else { 0 }
            $statusMessage = "Backing up Item $currentItemIndex of $totalItems"
            $UpdateProgress.Invoke($statusMessage, $percentComplete, "Backing up network drives...")
            Write-Host "Processing LOCAL Network Drives backup..."
            try {
                Get-WmiObject -Class Win32_MappedLogicalDisk -ErrorAction Stop |
                    Select-Object Name, ProviderName |
                    Export-Csv -Path (Join-Path $BackupRootPath "Drives.csv") -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
                Write-Host "Local network drives backed up successfully." -ForegroundColor Green
            } catch { Write-Warning "Failed to backup local network drives: $($_.Exception.Message)" }
        }

        # Backup Printers (Locally)
        if ($BackupPrinters) {
            $currentItemIndex++
            $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 100) } else { 0 }
            $statusMessage = "Backing up Item $currentItemIndex of $totalItems"
            $UpdateProgress.Invoke($statusMessage, $percentComplete, "Backing up printers...")
            Write-Host "Processing LOCAL Printers backup..."
            try {
                Get-WmiObject -Class Win32_Printer -Filter "Local = False" -ErrorAction Stop |
                    Select-Object -ExpandProperty Name |
                    Set-Content -Path (Join-Path $BackupRootPath "Printers.txt") -Encoding UTF8 -ErrorAction Stop
                 Write-Host "Local printers backed up successfully." -ForegroundColor Green
            } catch { Write-Warning "Failed to backup local printers: $($_.Exception.Message)" }
        }

        $UpdateProgress.Invoke("Backup Complete", 100, "Successfully backed up to: $BackupRootPath")
        Write-Host "--- LOCAL Backup Operation Finished ---"
        return $true # Indicate success

    } catch {
        $errorMessage = "LOCAL Backup Operation Failed: $($_.Exception.Message)"
        Write-Error $errorMessage
        $UpdateProgress.Invoke("Backup Failed", -1, $errorMessage) # Use -1 percent for error state
        return $false # Indicate failure
    }
}

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

    Write-Host "--- Starting LOCAL Restore Operation from '$BackupRootPath' ---"
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

            # Use the paths directly from the log for local restore
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

        # Restore Network Drives (Locally)
        if ($RestoreNetworkDrives) {
            $currentItemIndex++
            $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 100) } else { 0 }
            $statusMessage = "Restoring Item $currentItemIndex of $totalItems"
            $UpdateProgress.Invoke($statusMessage, $percentComplete, "Restoring network drives...")
            Write-Host "Processing LOCAL Network Drives restore..."
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

        # Restore Printers (Locally)
        if ($RestorePrinters) {
            $currentItemIndex++
            $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 100) } else { 0 }
            $statusMessage = "Restoring Item $currentItemIndex of $totalItems"
            $UpdateProgress.Invoke($statusMessage, $percentComplete, "Restoring printers...")
            Write-Host "Processing LOCAL Printers restore..."
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
        Write-Host "--- LOCAL Restore Operation Finished ---"
        return $true # Indicate success

    } catch {
        $errorMessage = "LOCAL Restore Operation Failed: $($_.Exception.Message)"
        Write-Error $errorMessage
        $UpdateProgress.Invoke("Restore Failed", -1, $errorMessage) # Use -1 percent for error state
        return $false # Indicate failure
    }
}

# --- Background Job Starter (For Local Restore Mode Only) ---
function Start-BackgroundUpdateJob {
    Write-Host "Initiating background system updates job (GPUpdate/SCCM Actions)..." -ForegroundColor Yellow
    $script:updateJob = Start-Job -Name "BackgroundUpdates" -ScriptBlock {

        # --- Function Definitions INSIDE Job Scope ---
        function Set-GPupdate {
            Write-Host "JOB: Initiating Group Policy update..." -ForegroundColor Cyan
            try {
                $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c gpupdate /force" -WorkingDirectory 'C:\Windows\System32' -PassThru -Wait -ErrorAction Stop
                if ($process.ExitCode -eq 0) { Write-Host "JOB: Group Policy update completed successfully." -ForegroundColor Green }
                else { Write-Warning "JOB: Group Policy update process finished with exit code: $($process.ExitCode)." }
            } catch { Write-Error "JOB: Failed to start GPUpdate process: $($_.Exception.Message)" }
            Write-Host "JOB: Exiting Set-GPupdate function."
        }
        function Start-ConfigManagerActions {
             param()
             Write-Host "JOB: Entering Start-ConfigManagerActions function."
             $ccmExecPath = "C:\Windows\CCM\ccmexec.exe"; $wmiNamespace = "root\ccm"; $wmiClassName = "SMS_Client"; $scheduleMethodName = "TriggerSchedule"; $overallSuccess = $false; $cimAttemptedAndSucceeded = $false
             $scheduleActions = @( @{ID = '{00000000-0000-0000-0000-000000000021}'; Name = 'Machine Policy Retrieval & Evaluation Cycle'}, @{ID = '{00000000-0000-0000-0000-000000000022}'; Name = 'Machine Policy Evaluation Cycle'}, @{ID = '{00000000-0000-0000-0000-000000000001}'; Name = 'Hardware Inventory Cycle'}, @{ID = '{00000000-0000-0000-0000-000000000002}'; Name = 'Software Inventory Cycle'}, @{ID = '{00000000-0000-0000-0000-000000000113}'; Name = 'Software Updates Scan Cycle'}, @{ID = '{00000000-0000-0000-0000-000000000108}'; Name = 'Software Updates Assignments Evaluation Cycle'} )
             Write-Host "JOB: Defined $($scheduleActions.Count) CM actions to trigger."
             $ccmService = Get-Service -Name CcmExec -ErrorAction SilentlyContinue
             if (-not $ccmService) { Write-Warning "JOB: CM service (CcmExec) not found. Skipping."; return $false }
             elseif ($ccmService.Status -ne 'Running') { Write-Warning "JOB: CM service (CcmExec) is not running (Status: $($ccmService.Status)). Skipping."; return $false }
             else { Write-Host "JOB: CM service (CcmExec) found and running." }
             Write-Host "JOB: Attempting Method 1: Triggering via CIM ($wmiNamespace -> $wmiClassName)..."
             $cimMethodSuccess = $true
             try {
                 Write-Host "JOB: Attempting to invoke '$scheduleMethodName' on '$wmiClassName' in '$wmiNamespace'."
                 foreach ($action in $scheduleActions) {
                    Write-Host "JOB:   Triggering $($action.Name) (ID: $($action.ID)) via CIM."
                    try { Invoke-CimMethod -Namespace $wmiNamespace -ClassName $wmiClassName -MethodName $scheduleMethodName -Arguments @{sScheduleID = $action.ID} -ErrorAction Stop; Write-Host "JOB:     $($action.Name) triggered successfully via CIM." }
                    catch { Write-Warning "JOB:     Failed to trigger $($action.Name) via CIM: $($_.Exception.Message)"; $cimMethodSuccess = $false }
                 }
                 if ($cimMethodSuccess) { $cimAttemptedAndSucceeded = $true; $overallSuccess = $true; Write-Host "JOB: All actions successfully triggered via CIM." -ForegroundColor Green }
                 else { Write-Warning "JOB: One or more actions failed to trigger via CIM." }
             } catch { Write-Error "JOB: An unexpected error occurred during CIM attempt: $($_.Exception.Message)"; $cimMethodSuccess = $false }
             if (-not $cimAttemptedAndSucceeded) {
                 Write-Host "JOB: CIM failed/unavailable. Attempting Method 2: Fallback via ccmexec.exe..."
                 if (Test-Path -Path $ccmExecPath -PathType Leaf) {
                     Write-Host "JOB: Found $ccmExecPath."
                     $execMethodSuccess = $true
                     foreach ($action in $scheduleActions) {
                         Write-Host "JOB:   Triggering $($action.Name) (ID: $($action.ID)) via ccmexec.exe."
                         try { $process = Start-Process -FilePath $ccmExecPath -ArgumentList "-TriggerSchedule $($action.ID)" -NoNewWindow -PassThru -Wait -ErrorAction Stop; if ($process.ExitCode -ne 0) { Write-Warning "JOB:     $($action.Name) via ccmexec.exe finished with exit code $($process.ExitCode). (Requires elevation)" } else { Write-Host "JOB:     $($action.Name) triggered via ccmexec.exe (Exit Code 0)." } }
                         catch { Write-Warning "JOB:     Failed to execute ccmexec.exe for $($action.Name): $($_.Exception.Message)"; $execMethodSuccess = $false }
                     }
                     if ($execMethodSuccess) { $overallSuccess = $true; Write-Host "JOB: Finished attempting actions via ccmexec.exe." -ForegroundColor Green }
                     else { Write-Warning "JOB: One or more actions failed to execute via ccmexec.exe (likely requires elevation)." }
                 } else { Write-Warning "JOB: Fallback executable not found at $ccmExecPath." }
             }
             if ($overallSuccess) { Write-Host "JOB: CM actions attempt finished successfully." -ForegroundColor Green }
             else { Write-Warning "JOB: CM actions attempt finished, but could not be confirmed as fully successful (check for elevation)." }
             Write-Host "JOB: Exiting Start-ConfigManagerActions function."
             return $overallSuccess
        }
        # Execute functions
        Set-GPupdate
        Start-ConfigManagerActions
        Write-Host "JOB: Background updates finished."
    } # End of Start-Job ScriptBlock
    Write-Host "Background update job started (ID: $($script:updateJob.Id)). Output will be shown later." -ForegroundColor Yellow
}

# --- Data Gathering Functions (Modified for Remote) ---
function Get-BackupPaths {
    [CmdletBinding()]
    param (
        [string]$RemoteDriveLetter, # e.g., "X"
        [string]$RemoteUsername     # e.g., "jdoe"
    )

    # Define path templates relative to a user profile
    $pathTemplates = @{
        "Signatures" = 'AppData\Roaming\Microsoft\Signatures';
        "QuickAccess" = 'AppData\Roaming\Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms';
        "StickyNotesLegacy" = 'AppData\Roaming\Microsoft\Sticky Notes\StickyNotes.snt';
        "StickyNotesModernDB" = 'AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\plum.sqlite';
        "GoogleEarthPlaces" = 'AppData\Roaming\google\googleearth\myplaces.kml'; # Note: AppData\LocalLow might be used too
        "ChromeBookmarks" = 'AppData\Local\Google\Chrome\User Data\Default\Bookmarks';
        "EdgeBookmarks" = 'AppData\Local\Microsoft\Edge\User Data\Default\Bookmarks';
    }

    $result = [System.Collections.Generic.List[PSCustomObject]]::new()
    $sourceBasePath = $null
    $isRemote = $false

    if (-not [string]::IsNullOrEmpty($RemoteDriveLetter) -and -not [string]::IsNullOrEmpty($RemoteUsername)) {
        # Construct remote base path
        $sourceBasePath = Join-Path -Path "$($RemoteDriveLetter):\" -ChildPath "Users\$RemoteUsername"
        $isRemote = $true
        Write-Host "Getting paths from REMOTE source: $sourceBasePath"
    } else {
        # Use local environment variables
        Write-Host "Getting paths from LOCAL source."
    }

    foreach ($key in $pathTemplates.Keys) {
        $relativePath = $pathTemplates[$key]
        $fullPath = $null

        if ($isRemote) {
            # Construct full remote path
            $fullPath = Join-Path -Path $sourceBasePath -ChildPath $relativePath
        } else {
            # Resolve path locally using environment variables implicitly contained in the template structure
            $localPathAttempt = $relativePath -replace 'AppData\\Roaming', $env:APPDATA `
                                             -replace 'AppData\\Local', $env:LOCALAPPDATA
            $fullPath = try { $ExecutionContext.InvokeCommand.ExpandString($localPathAttempt) } catch { $null }
        }

        if ($fullPath -and (Test-Path -LiteralPath $fullPath)) {
            $item = Get-Item -LiteralPath $fullPath -ErrorAction SilentlyContinue
            if ($item) {
                $result.Add([PSCustomObject]@{
                    Name = $item.Name;
                    Path = $item.FullName; # Source Path (Remote or Local)
                    Type = if ($item.PSIsContainer) { "Folder" } else { "File" };
                    RelativeUserProfilePath = $relativePath # Store the path relative to user profile for constructing destination
                })
            } else { Write-Host "Path resolved but Get-Item failed: $fullPath" }
        } else { Write-Host "Path not found or resolution failed: $fullPath (Derived from template key: $key)" }
    }
    return $result
}

function Get-UserPaths {
    [CmdletBinding()]
    param (
        [string]$RemoteDriveLetter, # e.g., "X"
        [string]$RemoteUsername     # e.g., "jdoe"
    )
    $folderNames = @( "Downloads", "Pictures", "Videos", "Desktop", "Documents" ) # Added Desktop/Documents
    $result = [System.Collections.Generic.List[PSCustomObject]]::new()
    $sourceBasePath = $null
    $isRemote = $false

    if (-not [string]::IsNullOrEmpty($RemoteDriveLetter) -and -not [string]::IsNullOrEmpty($RemoteUsername)) {
        $sourceBasePath = Join-Path -Path "$($RemoteDriveLetter):\" -ChildPath "Users\$RemoteUsername"
        $isRemote = $true
        Write-Host "Getting user folders from REMOTE source: $sourceBasePath"
    } else {
        $sourceBasePath = $env:USERPROFILE
        Write-Host "Getting user folders from LOCAL source: $sourceBasePath"
    }

    foreach ($folderName in $folderNames) {
        $fullPath = Join-Path -Path $sourceBasePath -ChildPath $folderName
        if (Test-Path -LiteralPath $fullPath -PathType Container) {
            $item = Get-Item -LiteralPath $fullPath -ErrorAction SilentlyContinue
             if ($item) {
                 $result.Add([PSCustomObject]@{
                     Name = $item.Name;
                     Path = $item.FullName; # Source Path (Remote or Local)
                     Type = "Folder";
                     RelativeUserProfilePath = $folderName # Store relative path for constructing destination
                 })
             } else { Write-Host "User folder resolved but Get-Item failed: $fullPath" }
        } else { Write-Host "User folder not found or not a folder: $fullPath" }
    }
    return $result
}

# --- Remote Data Capture Functions ---

function Get-RemoteMappedDrives {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    Write-Host "Attempting to get mapped drives from remote computer: $ComputerName"
    try {
        # Requires WinRM and permissions
        $remoteDrives = Get-CimInstance -ClassName Win32_MappedLogicalDisk -ComputerName $ComputerName -ErrorAction Stop
        Write-Host "Successfully retrieved $($remoteDrives.Count) mapped drive entries from $ComputerName."
        return $remoteDrives | Select-Object Name, ProviderName
    } catch {
        # --- FIX APPLIED HERE ---
        Write-Warning "Failed to get mapped drives from ${ComputerName}: $($_.Exception.Message)"
        return $null
    }
}

function Get-RemoteNetworkPrinters {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    Write-Host "Attempting to get network printers from remote computer: $ComputerName"
    try {
        # Requires WinRM and permissions
        $remotePrinters = Get-CimInstance -ClassName Win32_Printer -Filter "Local = False AND Network = True" -ComputerName $ComputerName -ErrorAction Stop
        Write-Host "Successfully retrieved $($remotePrinters.Count) network printer entries from $ComputerName."
        return $remotePrinters | Select-Object -ExpandProperty Name
    } catch {
        # --- FIX APPLIED HERE ---
        Write-Warning "Failed to get network printers from ${ComputerName}: $($_.Exception.Message)"
        return $null
    }
}

# --- Core Logic Function (Remote Transfer) ---

function Invoke-RemoteTransferOperation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$RemoteDeviceName, # For logging and potentially remote WMI

        [Parameter(Mandatory=$true)]
        [string]$MappedDriveLetter, # e.g., "X"

        [Parameter(Mandatory=$true)]
        [string]$LocalBackupPath, # e.g., C:\LocalData\Backup_User_Timestamp

        [Parameter(Mandatory=$false)]
        [scriptblock]$ProgressAction
    )

    Write-Host "--- Starting REMOTE Transfer Operation from '$RemoteDeviceName' (via $MappedDriveLetter): ---"
    $UpdateProgress = { if ($ProgressAction) { $ProgressAction.Invoke($args[0], $args[1], $args[2]) } else { Write-Host "$($args[0]) ($($args[1])%) - $($args[2])" } }
    $overallSuccess = $true # Assume success, set to false on critical errors
    $transferLog = [System.Collections.Generic.List[object]]::new() # Log successful transfers

    try {
        $UpdateProgress.Invoke("Initializing Transfer", 0, "Preparing transfer environment...")

        # *** ASSUMPTION: Remote username is the same as local username ***
        $remoteUsername = $env:USERNAME
        Write-Warning "Assuming remote username is the same as local: '$remoteUsername'. If incorrect, paths will be wrong."

        # Create the local backup directory for the *copy* of transferred files
        if (-not (Test-Path $LocalBackupPath)) {
            New-Item -Path $LocalBackupPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Write-Host "Local backup directory created: $LocalBackupPath"
        }
        # Create transfer log file within the backup path
        $transferLogPath = Join-Path $LocalBackupPath "TransferLog.csv"
        "RemoteSourcePath,LocalDestinationPath,Status,Notes" | Set-Content -Path $transferLogPath -Encoding UTF8

        # 1. Get Remote Paths to Transfer
        $UpdateProgress.Invoke("Gathering Paths", 5, "Identifying remote files/folders...")
        $itemsToTransfer = [System.Collections.Generic.List[PSCustomObject]]::new()
        try { Get-BackupPaths -RemoteDriveLetter $MappedDriveLetter -RemoteUsername $remoteUsername | ForEach-Object { $itemsToTransfer.Add($_) } } catch { Write-Warning "Error getting standard backup paths from remote: $($_.Exception.Message)"}
        # Optionally add User Folders (Downloads, Desktop etc.)
        try { Get-UserPaths -RemoteDriveLetter $MappedDriveLetter -RemoteUsername $remoteUsername | ForEach-Object { $itemsToTransfer.Add($_) } } catch { Write-Warning "Error getting user folders from remote: $($_.Exception.Message)"}

        if ($itemsToTransfer.Count -eq 0) {
            Write-Warning "No files or folders found to transfer based on defined paths for user '$remoteUsername' on '$RemoteDeviceName'."
        } else {
             Write-Host "Identified $($itemsToTransfer.Count) items/folders to attempt transfer from '$RemoteDeviceName'."
        }

        # 2. Transfer Files/Folders
        $currentItemIndex = 0
        $totalItemsEstimate = $itemsToTransfer.Count # Add drive/printer counts later if desired for progress
        Write-Host "Starting file/folder transfer..."

        foreach ($item in $itemsToTransfer) {
            $currentItemIndex++
            $percentComplete = if ($totalItemsEstimate -gt 0) { [int](($currentItemIndex / $totalItemsEstimate) * 50) } else { 0 } # File transfer is ~50% of work
            $statusMessage = "Transferring Item $currentItemIndex of $totalItemsEstimate"
            $UpdateProgress.Invoke($statusMessage, $percentComplete, "Transferring: $($item.Name)")

            $remoteSourcePath = $item.Path
            # Construct LOCAL destination path using the relative path and local env vars
            $localDestinationPath = $null
            $relativePath = $item.RelativeUserProfilePath
            if ($relativePath) {
                 $localPathAttempt = $relativePath -replace 'AppData\\Roaming', $env:APPDATA `
                                                  -replace 'AppData\\Local', $env:LOCALAPPDATA
                 if ($localPathAttempt -eq $relativePath) { # Handle top-level folders like Downloads, Desktop
                     $localPathAttempt = Join-Path $env:USERPROFILE $relativePath
                 }
                 $localDestinationPath = try { $ExecutionContext.InvokeCommand.ExpandString($localPathAttempt) } catch { $null }
            }

            if (-not $localDestinationPath) {
                Write-Warning "Could not determine local destination path for remote source '$remoteSourcePath'. Skipping."
                "'$remoteSourcePath','N/A','Skipped','Could not determine local destination'" | Add-Content -Path $transferLogPath -Encoding UTF8
                continue
            }

            Write-Host "Attempting transfer:"
            Write-Host "  Remote Source: $remoteSourcePath"
            Write-Host "  Local Dest.:   $localDestinationPath"

            $itemSuccess = $false
            $errorNote = ""
            try {
                $targetDir = if ($item.Type -eq "Folder") { $localDestinationPath } else { [System.IO.Path]::GetDirectoryName($localDestinationPath) }
                if (-not (Test-Path $targetDir)) {
                    Write-Host "  Creating local directory: $targetDir"
                    New-Item -Path $targetDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
                }
                Copy-Item -LiteralPath $remoteSourcePath -Destination $localDestinationPath -Recurse -Force -ErrorAction Stop
                Write-Host "  Successfully transferred '$($item.Name)' to '$localDestinationPath'."
                $itemSuccess = $true
                $transferLog.Add([PSCustomObject]@{ LocalPath = $localDestinationPath; Type = $item.Type; Name = $item.Name }) | Out-Null
            } catch {
                $errorNote = $_.Exception.Message -replace '[\r\n]',' '
                Write-Warning "  Failed to transfer '$($item.Name)': $errorNote"
                $overallSuccess = $false
            }
            $status = if ($itemSuccess) { "Success" } else { "Failed" }
            "'$remoteSourcePath','$localDestinationPath','$status','$errorNote'" | Add-Content -Path $transferLogPath -Encoding UTF8
        } # End foreach item

        # 3. Capture and Recreate Network Drives
        $UpdateProgress.Invoke("Transferring Settings", 60, "Capturing remote network drives...")
        $remoteDrives = Get-RemoteMappedDrives -ComputerName $RemoteDeviceName
        $localDrivesBackupPath = Join-Path $LocalBackupPath "RemoteDrives.csv"
        if ($remoteDrives) {
            Write-Host "Attempting to recreate $($remoteDrives.Count) network drives locally..."
            $remoteDrives | Export-Csv -Path $localDrivesBackupPath -NoTypeInformation -Encoding UTF8
            "'$RemoteDeviceName Drives','N/A','Success','Backed up drive list'" | Add-Content -Path $transferLogPath -Encoding UTF8
            foreach ($drive in $remoteDrives) {
                $driveLetter = $drive.Name.TrimEnd(':')
                $networkPath = $drive.ProviderName
                if ($driveLetter -match '^[A-Z]$' -and $networkPath -match '^\\\\') {
                    if (-not (Test-Path -LiteralPath "$($driveLetter):")) {
                        try {
                            Write-Host "  Mapping local $driveLetter to $networkPath"
                            New-PSDrive -Name $driveLetter -PSProvider FileSystem -Root $networkPath -Persist -Scope Global -ErrorAction Stop
                            "'$networkPath','$($driveLetter):','Success','Mapped drive locally'" | Add-Content -Path $transferLogPath -Encoding UTF8
                        } catch {
                            $errorNote = $_.Exception.Message -replace '[\r\n]',' '; Write-Warning "  Failed to map local drive $driveLetter`: $errorNote"
                            "'$networkPath','$($driveLetter):','Failed','$errorNote'" | Add-Content -Path $transferLogPath -Encoding UTF8; $overallSuccess = $false
                         }
                    } else {
                        Write-Host "  Local drive $driveLetter already exists, skipping."
                        "'$networkPath','$($driveLetter):','Skipped','Local drive letter exists'" | Add-Content -Path $transferLogPath -Encoding UTF8
                    }
                } else {
                    Write-Warning "  Skipping invalid remote drive mapping: Name='$($drive.Name)', Provider='$networkPath'"
                    "'$($drive.Name)','$($networkPath)','Skipped','Invalid remote mapping data'" | Add-Content -Path $transferLogPath -Encoding UTF8
                }
            }
        } else {
             Write-Warning "Could not retrieve mapped drives from $RemoteDeviceName."
             "'$RemoteDeviceName Drives','N/A','Failed','Could not retrieve remote list'" | Add-Content -Path $transferLogPath -Encoding UTF8; $overallSuccess = $false
        }

        # 4. Capture and Recreate Network Printers
        $UpdateProgress.Invoke("Transferring Settings", 75, "Capturing remote network printers...")
        $remotePrinters = Get-RemoteNetworkPrinters -ComputerName $RemoteDeviceName
        $localPrintersBackupPath = Join-Path $LocalBackupPath "RemotePrinters.txt"
        if ($remotePrinters) {
            Write-Host "Attempting to recreate $($remotePrinters.Count) network printers locally..."
            $remotePrinters | Set-Content -Path $localPrintersBackupPath -Encoding UTF8
            "'$RemoteDeviceName Printers','N/A','Success','Backed up printer list'" | Add-Content -Path $transferLogPath -Encoding UTF8
            try {
                $wsNet = New-Object -ComObject WScript.Network
                foreach ($printerPath in $remotePrinters) {
                    if (-not ([string]::IsNullOrWhiteSpace($printerPath)) -and $printerPath -match '^\\\\') {
                        Write-Host "  Attempting to add local printer connection: $printerPath"
                        try {
                             $wsNet.AddWindowsPrinterConnection($printerPath); Write-Host "    Added printer connection (or it already existed)."
                             "'$printerPath','Local Printer','Success','Added printer connection'" | Add-Content -Path $transferLogPath -Encoding UTF8
                        } catch {
                            $errorNote = $_.Exception.Message -replace '[\r\n]',' '; Write-Warning "    Failed to add printer '$printerPath': $errorNote"
                            "'$printerPath','Local Printer','Failed','$errorNote'" | Add-Content -Path $transferLogPath -Encoding UTF8; $overallSuccess = $false
                        }
                    } else {
                        Write-Warning "  Skipping invalid remote printer path: '$printerPath'"
                        "'$printerPath','Local Printer','Skipped','Invalid remote printer path'" | Add-Content -Path $transferLogPath -Encoding UTF8
                    }
                }
            } catch {
                $errorNote = $_.Exception.Message -replace '[\r\n]',' '; Write-Warning "Error recreating printers locally: $errorNote"
                "'Printers Overall','Local Printer','Failed','$errorNote'" | Add-Content -Path $transferLogPath -Encoding UTF8; $overallSuccess = $false
            }
        } else {
            Write-Warning "Could not retrieve network printers from $RemoteDeviceName."
            "'$RemoteDeviceName Printers','N/A','Failed','Could not retrieve remote list'" | Add-Content -Path $transferLogPath -Encoding UTF8; $overallSuccess = $false
        }

        # 5. Create Local Backup of Transferred Files
        $UpdateProgress.Invoke("Creating Local Backup", 90, "Copying transferred files to backup folder...")
        Write-Host "Creating local backup copy of successfully transferred items in '$LocalBackupPath'..."
        if ($transferLog.Count -gt 0) {
            foreach($logEntry in $transferLog) {
                $localSourcePath = $logEntry.LocalPath
                $backupDestPath = Join-Path $LocalBackupPath $logEntry.Name

                if (Test-Path -LiteralPath $localSourcePath) {
                    Write-Host "  Backing up: $($logEntry.Name)"
                    try {
                        if ($logEntry.Type -eq "Folder") { Copy-Item -LiteralPath $localSourcePath -Destination $backupDestPath -Recurse -Force -ErrorAction Stop }
                        else { $backupDir = [System.IO.Path]::GetDirectoryName($backupDestPath); if (-not (Test-Path $backupDir)) { New-Item -Path $backupDir -ItemType Directory -Force -ErrorAction Stop | Out-Null }; Copy-Item -LiteralPath $localSourcePath -Destination $backupDestPath -Force -ErrorAction Stop }
                    } catch { Write-Warning "  Failed to backup transferred item '$($logEntry.Name)': $($_.Exception.Message)" }
                } else { Write-Warning "  Skipping backup for '$($logEntry.Name)' as local source path '$localSourcePath' not found (transfer might have failed)." }
            }
        } else { Write-Host "No items were successfully transferred to create a local backup copy." }

        $UpdateProgress.Invoke("Transfer Complete", 100, "Remote transfer finished. See log: $transferLogPath")
        Write-Host "--- REMOTE Transfer Operation Finished ---"
        return $overallSuccess

    } catch {
        $errorMessage = "REMOTE Transfer Operation Failed Critically: $($_.Exception.Message)"
        Write-Error $errorMessage
        $UpdateProgress.Invoke("Transfer Failed", -1, $errorMessage)
        "'Overall Operation','N/A','Critical Failure','$($errorMessage -replace '[\r\n]',' ')'" | Add-Content -Path $transferLogPath -Encoding UTF8 -ErrorAction SilentlyContinue
        return $false
    }
}

# --- GUI Functions ---
function Show-ModeDialog {
    Write-Host "Entering Show-ModeDialog function."
    $form = New-Object System.Windows.Forms.Form; $form.Text = "Select Operation"; $form.Size = New-Object System.Drawing.Size(400, 150); $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen; $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog; $form.MaximizeBox = $false; $form.MinimizeBox = $false; $form.HelpButton = $false
    $btnBackup = New-Object System.Windows.Forms.Button; $btnBackup.Location = New-Object System.Drawing.Point(30, 40); $btnBackup.Size = New-Object System.Drawing.Size(80, 30); $btnBackup.Text = "Backup"; $btnBackup.DialogResult = [System.Windows.Forms.DialogResult]::Yes; $form.Controls.Add($btnBackup)
    $btnRestore = New-Object System.Windows.Forms.Button; $btnRestore.Location = New-Object System.Drawing.Point(140, 40); $btnRestore.Size = New-Object System.Drawing.Size(80, 30); $btnRestore.Text = "Restore"; $btnRestore.DialogResult = [System.Windows.Forms.DialogResult]::No; $form.Controls.Add($btnRestore)
    $btnExpress = New-Object System.Windows.Forms.Button; $btnExpress.Location = New-Object System.Drawing.Point(250, 40); $btnExpress.Size = New-Object System.Drawing.Size(80, 30); $btnExpress.Text = "Express"; $btnExpress.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Controls.Add($btnExpress)
    $form.AcceptButton = $btnExpress; $form.CancelButton = $btnRestore
    Write-Host "Showing mode selection dialog."; $result = $form.ShowDialog(); $form.Dispose(); Write-Host "Mode selection dialog closed with result: $result"
    $selectedMode = switch ($result) { ([System.Windows.Forms.DialogResult]::Yes) { 'Backup' } ([System.Windows.Forms.DialogResult]::No) { 'Restore' } ([System.Windows.Forms.DialogResult]::OK) { 'Express' } Default { 'Cancel' } }
    Write-Host "Determined mode: $selectedMode" -ForegroundColor Cyan; return $selectedMode
}

function Show-MainWindow {
    param( [Parameter(Mandatory=$true)] [bool]$IsBackup )
    $modeString = if ($IsBackup) { 'Backup' } else { 'Restore' }; Write-Host "Entering Show-MainWindow function. Mode: $modeString"
    [xml]$XAML = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="User Data Backup/Restore Tool" Width="800" Height="650" WindowStartupLocation="CenterScreen">
    <Grid>
        <Label Content="Location:" Margin="10,10,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/><TextBox Name="txtSaveLoc" Width="400" Height="30" Margin="10,40,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" IsReadOnly="True"/><Button Name="btnBrowse" Content="Browse" Width="60" Height="30" Margin="420,40,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/><Label Name="lblMode" Content="" Margin="500,10,10,0" HorizontalAlignment="Right" VerticalAlignment="Top" FontWeight="Bold"/><Label Name="lblFreeSpace" Content="Free Space: -" Margin="500,40,10,0" HorizontalAlignment="Right" VerticalAlignment="Top"/><Label Name="lblRequiredSpace" Content="Required Space: -" Margin="500,65,10,0" HorizontalAlignment="Right" VerticalAlignment="Top"/><Label Name="lblStatus" Content="Ready" Margin="10,0,10,10" HorizontalAlignment="Center" VerticalAlignment="Bottom" FontStyle="Italic"/><Label Content="Files/Folders to Process:" Margin="10,90,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <ListView Name="lvwFiles" Margin="10,120,200,140" SelectionMode="Extended"><ListView.View><GridView><GridViewColumn Width="30"><GridViewColumn.CellTemplate><DataTemplate><CheckBox IsChecked="{Binding IsSelected, Mode=TwoWay}" /></DataTemplate></GridViewColumn.CellTemplate></GridViewColumn><GridViewColumn Header="Name" DisplayMemberBinding="{Binding Name}" Width="180"/><GridViewColumn Header="Type" DisplayMemberBinding="{Binding Type}" Width="70"/><GridViewColumn Header="Path" DisplayMemberBinding="{Binding Path}" Width="280"/></GridView></ListView.View></ListView>
        <StackPanel Margin="0,120,10,0" HorizontalAlignment="Right" Width="180"><Button Name="btnAddFile" Content="Add File" Width="120" Height="30" Margin="0,0,0,10"/><Button Name="btnAddFolder" Content="Add Folder" Width="120" Height="30" Margin="0,0,0,10"/><Button Name="btnAddBAUPaths" Content="Add User Folders" Width="120" Height="30" Margin="0,0,0,10"/><Button Name="btnRemove" Content="Remove Selected" Width="120" Height="30" Margin="0,0,0,20"/><CheckBox Name="chkNetwork" Content="Network Drives" IsChecked="True" Margin="5,0,0,5"/><CheckBox Name="chkPrinters" Content="Printers" IsChecked="True" Margin="5,0,0,5"/></StackPanel>
        <ProgressBar Name="prgProgress" Height="20" Margin="10,0,10,60" VerticalAlignment="Bottom"/><TextBlock Name="txtProgress" Text="" Margin="10,0,10,85" VerticalAlignment="Bottom" TextWrapping="Wrap"/><Button Name="btnStart" Content="Start" Width="100" Height="30" Margin="10,0,0,20" VerticalAlignment="Bottom" HorizontalAlignment="Left" IsDefault="True"/>
    </Grid>
</Window>
'@
    try {
        Write-Host "Parsing XAML..."; $reader = New-Object System.Xml.XmlNodeReader $XAML; $window = [Windows.Markup.XamlReader]::Load($reader); $window.DataContext = [PSCustomObject]@{ IsRestoreMode = (-not $IsBackup) }
        Write-Host "Finding controls..."; $controls = @{}; $window.FindName('txtSaveLoc') | %{$controls['txtSaveLoc']=$_}; $window.FindName('btnBrowse') | %{$controls['btnBrowse']=$_}; $window.FindName('btnStart') | %{$controls['btnStart']=$_}; $window.FindName('lblMode') | %{$controls['lblMode']=$_}; $window.FindName('lblStatus') | %{$controls['lblStatus']=$_}; $window.FindName('lvwFiles') | %{$controls['lvwFiles']=$_}; $window.FindName('btnAddFile') | %{$controls['btnAddFile']=$_}; $window.FindName('btnAddFolder') | %{$controls['btnAddFolder']=$_}; $window.FindName('btnRemove') | %{$controls['btnRemove']=$_}; $window.FindName('chkNetwork') | %{$controls['chkNetwork']=$_}; $window.FindName('chkPrinters') | %{$controls['chkPrinters']=$_}; $window.FindName('prgProgress') | %{$controls['prgProgress']=$_}; $window.FindName('txtProgress') | %{$controls['txtProgress']=$_}; $window.FindName('btnAddBAUPaths') | %{$controls['btnAddBAUPaths']=$_}; $window.FindName('lblFreeSpace') | %{$controls['lblFreeSpace']=$_}; $window.FindName('lblRequiredSpace') | %{$controls['lblRequiredSpace']=$_}

        $script:UpdateFreeSpaceLabel = { param($ControlsParam) $location = $ControlsParam.txtSaveLoc.Text; $freeSpaceString = "Free Space: N/A"; if (-not [string]::IsNullOrEmpty($location)) { try { $driveLetter = $null; if ($location -match '^[a-zA-Z]:\\') { $driveLetter = $location.Substring(0, 2) } elseif ($location -match '^\\\\[^\\]+\\[^\\]+') { $freeSpaceString = "Free Space: N/A (UNC)" }; if ($driveLetter) { $driveInfo = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$driveLetter'" -EA SilentlyContinue; if ($driveInfo -and $driveInfo.FreeSpace -ne $null) { $freeSpaceString = "Free Space: $(Format-Bytes $driveInfo.FreeSpace)" } else { Write-Warning "Could not get free space for drive $driveLetter" } } } catch { Write-Warning "Error getting free space for '$location': $($_.Exception.Message)" } }; $ControlsParam.lblFreeSpace.Content = $freeSpaceString; Write-Host "Updated Free Space Label: $freeSpaceString" }
        $script:UpdateRequiredSpaceLabel = { param($ControlsParam) $totalSize = 0L; $requiredSpaceString = "Required Space: Calculating..."; $ControlsParam.lblRequiredSpace.Content = $requiredSpaceString; $items = @($ControlsParam.lvwFiles.ItemsSource); if ($items -ne $null -and $items.Count -gt 0) { $checkedItems = $items | Where-Object { $_.IsSelected }; Write-Host "Calculating required space for $($checkedItems.Count) checked items..."; foreach ($item in $checkedItems) { if (Test-Path -LiteralPath $item.Path) { try { if ($item.Type -eq 'Folder') { $folderSize = (Get-ChildItem -LiteralPath $item.Path -Recurse -File -Force -EA SilentlyContinue | Measure-Object -Property Length -Sum -EA SilentlyContinue).Sum; if ($folderSize -ne $null) { $totalSize += $folderSize } } else { $fileSize = (Get-Item -LiteralPath $item.Path -Force -EA SilentlyContinue).Length; if ($fileSize -ne $null) { $totalSize += $fileSize } } } catch { Write-Warning "Error calculating size for '$($item.Path)': $($_.Exception.Message)" } } else { Write-Warning "Checked item path not found, skipping size calculation: $($item.Path)" } }; $requiredSpaceString = "Required Space: $(Format-Bytes $totalSize)" } else { $requiredSpaceString = "Required Space: 0 B" }; $ControlsParam.lblRequiredSpace.Dispatcher.Invoke({ $ControlsParam.lblRequiredSpace.Content = $requiredSpaceString }, [System.Windows.Threading.DispatcherPriority]::Background); Write-Host "Updated Required Space Label: $requiredSpaceString" }

        Write-Host "Initializing controls..."; $controls.lblMode.Content = if ($IsBackup) { "Mode: Backup" } else { "Mode: Restore" }; $controls.btnStart.Content = if ($IsBackup) { "Backup" } else { "Restore" }; $controls.btnAddBAUPaths.IsEnabled = $IsBackup
        if (-not (Test-Path $script:DefaultPath)) { try { New-Item -Path $script:DefaultPath -ItemType Directory -Force -EA Stop | Out-Null } catch { Write-Warning "Could not create default path: $script:DefaultPath."; $script:DefaultPath = $env:USERPROFILE } }; $controls.txtSaveLoc.Text = $script:DefaultPath; & $script:UpdateFreeSpaceLabel -ControlsParam $controls

        Write-Host "Loading initial items..."; $initialItemsList = [System.Collections.Generic.List[PSCustomObject]]::new()
        if ($IsBackup) { try { $paths = Get-BackupPaths; if ($paths) { $paths | % { $initialItemsList.Add(($_ | Add-Member -MemberType NoteProperty -Name 'IsSelected' -Value $true -PassThru)) } } } catch { Write-Error "Error calling Get-BackupPaths: $($_.Exception.Message)" } }
        else { $latestBackup = Get-ChildItem -Path $script:DefaultPath -Directory -Filter "Backup_*" | Sort-Object LastWriteTime -Descending | Select-Object -First 1; if ($latestBackup) { $controls.txtSaveLoc.Text = $latestBackup.FullName; & $script:UpdateFreeSpaceLabel -ControlsParam $controls; $backupItems = Get-ChildItem -Path $latestBackup.FullName | Where-Object { $_.Name -notmatch '^(FileList_.*\.csv|Drives\.csv|Printers\.txt|TransferLog\.csv)$' } | % { [PSCustomObject]@{ Name = $_.Name; Type = if ($_.PSIsContainer) { "Folder" } else { "File" }; Path = $_.FullName; IsSelected = $true } }; $backupItems | % { $initialItemsList.Add($_) } } else { $controls.lblStatus.Content = "Restore mode: No backups found in $script:DefaultPath. Please browse." } }
        $controls.lvwFiles.ItemsSource = $initialItemsList; & $script:UpdateRequiredSpaceLabel -ControlsParam $controls

        Write-Host "Assigning event handlers..."; $controls.btnBrowse.Add_Click({...}); $controls.btnAddFile.Add_Click({...}); $controls.btnAddFolder.Add_Click({...}); $controls.btnAddBAUPaths.Add_Click({...}); $controls.btnRemove.Add_Click({...}); $controls.lvwFiles.Add_PreviewMouseLeftButtonUp({...}) # Keep existing event handler logic

        $controls.btnStart.Add_Click({
            $modeStringLocal = if ($IsBackup) { 'Backup' } else { 'Restore' }; Write-Host "Start button clicked. Mode: $modeStringLocal"
            $location = $controls.txtSaveLoc.Text; if ([string]::IsNullOrEmpty($location) -or -not (Test-Path $location -PathType Container)) { [System.Windows.MessageBox]::Show("Please select a valid target directory first.", "Location Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning); return }
            Write-Host "Recalculating required space..."; & $script:UpdateRequiredSpaceLabel -ControlsParam $controls
            $itemsFromUI = [System.Collections.Generic.List[PSCustomObject]]::new(); $checkedItems = @($controls.lvwFiles.ItemsSource) | Where-Object { $_.IsSelected }; if ($checkedItems) { $checkedItems | % { $itemsFromUI.Add($_) } }
            $doNetwork = $controls.chkNetwork.IsChecked; $doPrinters = $controls.chkPrinters.IsChecked
            Write-Host "Disabling UI..."; $controls | % { if ($_.Value -is [System.Windows.Controls.Control]) { $_.Value.IsEnabled = $false } }; $window.Cursor = [System.Windows.Input.Cursors]::Wait; $controls.prgProgress.IsIndeterminate = $true; $controls.prgProgress.Value = 0; $controls.txtProgress.Text = "Initializing..."; $controls.lblStatus.Content = "Starting $modeStringLocal..."
            $uiProgressAction = { param($status, $percent, $details) $window.Dispatcher.InvokeAsync( [action]{ $controls.lblStatus.Content = $status; $controls.txtProgress.Text = $details; if ($percent -ge 0) { $controls.prgProgress.IsIndeterminate = $false; $controls.prgProgress.Value = $percent } else { $controls.prgProgress.IsIndeterminate = $false; $controls.prgProgress.Value = 0 } }, [System.Windows.Threading.DispatcherPriority]::Background ) | Out-Null }
            $success = $false
            try {
                if ($IsBackup) { $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"; $username = $env:USERNAME -replace '[^a-zA-Z0-9]', '_'; $backupRootPath = Join-Path $location "Backup_${username}_$timestamp"; if ($itemsFromUI.Count -eq 0) { throw "No items selected (checked) for backup." }; $success = Invoke-BackupOperation -BackupRootPath $backupRootPath -ItemsToBackup $itemsFromUI -BackupNetworkDrives $doNetwork -BackupPrinters $doPrinters -ProgressAction $uiProgressAction }
                else { $backupRootPath = $location; Write-Warning "UI Restore currently triggers restore of ALL items from log in the selected folder."; $success = Invoke-RestoreOperation -BackupRootPath $backupRootPath -RestoreAllFromLog -RestoreNetworkDrives $doNetwork -RestorePrinters $doPrinters -ProgressAction $uiProgressAction }
                if ($success) { Write-Host "Operation completed successfully (UI)." -ForegroundColor Green; $controls.lblStatus.Content = "Operation completed successfully."; [System.Windows.MessageBox]::Show("The $modeStringLocal operation completed successfully!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) }
                else { Write-Error "Operation failed (UI)."; $controls.lblStatus.Content = "Operation Failed!"; [System.Windows.MessageBox]::Show("The $modeStringLocal operation failed. Check details.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) }
            } catch { $errorMessage = "Operation Failed (Inner Catch): $($_.Exception.Message)"; Write-Error $errorMessage; $controls.lblStatus.Content = "Operation Failed!"; $controls.txtProgress.Text = $errorMessage; $controls.prgProgress.Value = 0; $controls.prgProgress.IsIndeterminate = $false; [System.Windows.MessageBox]::Show($errorMessage, "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) }
            finally { Write-Host "Operation finished (inner finally block). Re-enabling UI controls."; $controls | % { if ($_.Value -is [System.Windows.Controls.Control]) { if ($_.Key -eq 'btnAddBAUPaths') { $_.Value.IsEnabled = $IsBackup } else { $_.Value.IsEnabled = $true } } }; $window.Cursor = [System.Windows.Input.Cursors]::Arrow; Write-Host "Cursor reset." }
        })
        Write-Host "Showing main window."; $window.ShowDialog() | Out-Null; Write-Host "Main window closed."
    } catch { $errorMessage = "Failed to load main window: $($_.Exception.Message)"; Write-Error $errorMessage; Write-Host $errorMessage -ForegroundColor Red; try { [System.Windows.MessageBox]::Show($errorMessage, "Critical Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) } catch {} }
    finally { Write-Host "Exiting Show-MainWindow function."; Remove-Variable -Name UpdateFreeSpaceLabel -Scope Script -EA SilentlyContinue; Remove-Variable -Name UpdateRequiredSpaceLabel -Scope Script -EA SilentlyContinue }
}

# --- New Function to Encapsulate Express Mode Logic ---
function Execute-ExpressModeLogic {
    Write-Host "Executing Express Mode Logic..." -ForegroundColor Magenta

    # 1. Get Remote Device Name
    $RemoteDeviceResult = Show-InputBox -Prompt "Enter the Asset Tag or IP Address of the OLD device:" -Title "Remote Device"
    if ([string]::IsNullOrWhiteSpace($RemoteDeviceResult)) { Write-Warning "No remote device entered. Exiting Express mode."; [System.Windows.MessageBox]::Show("Remote device name/IP cannot be empty.", "Input Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning); return }
    Write-Host "Target remote device: $RemoteDeviceResult"

    # 2. Map Network Drive
    $mappedDrive = $null; $remoteCShare = "\\$RemoteDeviceResult\c$"; Write-Host "Attempting to map $remoteCShare to $($script:RemoteDriveLetter):"
    try { if (Test-Path -LiteralPath "$($script:RemoteDriveLetter):") { throw "Drive letter $($script:RemoteDriveLetter): is already in use." }; $mappedDrive = New-PSDrive -Name $script:RemoteDriveLetter -PSProvider FileSystem -Root $remoteCShare -ErrorAction Stop; Write-Host "Successfully mapped $remoteCShare to $($script:RemoteDriveLetter):" -ForegroundColor Green }
    catch { Write-Error "Failed to map network drive to $remoteCShare : $($_.Exception.Message)"; [System.Windows.MessageBox]::Show("Failed to map drive to '$remoteCShare'.`nError: $($_.Exception.Message)`n`nPlease ensure the device is online, accessible, and you have administrative permissions.", "Mapping Failed", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error); return }

    # 3. Define Local Backup Path for this transfer
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"; $username = $env:USERNAME -replace '[^a-zA-Z0-9]', '_'; $localBackupRootPath = Join-Path $script:DefaultPath "TransferBackup_${username}_from_${RemoteDeviceResult}_$timestamp"

    # Define Progress Action for Console
    $consoleProgressAction = { param($status, $percent, $details) Write-Host "[$($status) - $percent%]: $details" }

    # 4. Execute Remote Transfer Operation
    $transferSuccess = $false
    try { $transferSuccess = Invoke-RemoteTransferOperation -RemoteDeviceName $RemoteDeviceResult -MappedDriveLetter $script:RemoteDriveLetter -LocalBackupPath $localBackupRootPath -ProgressAction $consoleProgressAction }
    catch { Write-Error "A critical error occurred during the remote transfer process: $($_.Exception.Message)"; $transferSuccess = $false }
    finally { # 5. Unmap Network Drive (ALWAYS attempt this)
        if ($mappedDrive -ne $null) { Write-Host "Attempting to remove mapped drive $($script:RemoteDriveLetter):"; try { Remove-PSDrive -Name $script:RemoteDriveLetter -Force -ErrorAction Stop; Write-Host "Successfully removed mapped drive." -ForegroundColor Green } catch { Write-Warning "Failed to automatically remove mapped drive $($script:RemoteDriveLetter):. Please remove it manually. Error: $($_.Exception.Message)" } }
    }

    # 6. Final User Feedback
    if ($transferSuccess) { Write-Host "Express Transfer completed. Some items may have failed, check log: $($localBackupRootPath)\TransferLog.csv" -ForegroundColor Green; [System.Windows.MessageBox]::Show("Express Transfer completed.`nData transferred to local profile.`nBackup copy created in:`n$localBackupRootPath`n`nCheck TransferLog.csv inside the backup folder for details/errors.", "Express Transfer Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) }
    else { Write-Error "Express Transfer failed or encountered significant errors. Check console output and log: $($localBackupRootPath)\TransferLog.csv"; [System.Windows.MessageBox]::Show("Express Transfer failed or encountered significant errors.`nCheck console output and TransferLog.csv in `n$localBackupRootPath`n(if created) for details.", "Express Transfer Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) }
}

#endregion Functions

# --- Main Execution ---
Write-Host "--- Script Starting ---"
Clear-Variable -Name updateJob -Scope Script -ErrorAction SilentlyContinue

# Ensure Default Path Exists
if (-not (Test-Path $script:DefaultPath)) { Write-Host "Default path '$($script:DefaultPath)' not found. Attempting to create."; try { New-Item -Path $script:DefaultPath -ItemType Directory -Force -EA Stop | Out-Null; Write-Host "Default path created." } catch { Write-Warning "Could not create default path: $($script:DefaultPath). Express mode might fail." } }

try {
    # --- Check if script was relaunched for Elevated Express Mode ---
    if ($RunExpressElevated) {
        Write-Host "Script re-launched with elevation for Express Mode." -ForegroundColor Yellow
        Execute-ExpressModeLogic # Directly execute the Express logic
    } else {
        # --- Normal Script Start: Show Mode Dialog ---
        Write-Host "Calling Show-ModeDialog to determine operation mode."
        $script:selectedMode = Show-ModeDialog

        switch ($script:selectedMode) {
            'Backup' { Write-Host "Mode selected: Backup. Showing main window."; Show-MainWindow -IsBackup $true }
            'Restore' { Write-Host "Mode selected: Restore. Starting background updates and showing main window."; Start-BackgroundUpdateJob; Show-MainWindow -IsBackup $false }
            'Express' {
                Write-Host "Mode selected: Express. Checking elevation..."
                if (-not (Test-IsAdmin)) {
                    Write-Host "Elevation required for Express mode. Attempting to relaunch as Administrator..."
                    $msgResult = [System.Windows.MessageBox]::Show("Express mode requires Administrator privileges to map drives and access remote data.`n`nThe script will now attempt to restart itself with elevation.`n`nPlease approve the UAC prompt.", "Elevation Required", [System.Windows.MessageBoxButton]::OKCancel, [System.Windows.MessageBoxImage]::Information)
                    if ($msgResult -eq [System.Windows.MessageBoxResult]::OK) {
                        $scriptPath = $PSCommandPath; $arguments = "-ExecutionPolicy Bypass -File `"$scriptPath`" -RunExpressElevated"
                        try { Start-Process powershell.exe -ArgumentList $arguments -Verb RunAs -ErrorAction Stop; Write-Host "Elevated process started. Exiting current instance."; Exit }
                        catch { Write-Error "Failed to relaunch script with elevation: $($_.Exception.Message)"; [System.Windows.MessageBox]::Show("Failed to automatically restart with elevation.`n`nPlease right-click the script and choose 'Run as administrator'.`n`nError: $($_.Exception.Message)", "Elevation Failed", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error); Exit }
                    } else { Write-Host "User cancelled elevation request. Exiting Express mode." }
                } else { Write-Host "Already running as Administrator. Proceeding with Express mode."; Execute-ExpressModeLogic }
            }
            'Cancel' { Write-Host "Operation cancelled by user." }
            Default { Write-Warning "Invalid mode returned from dialog: $script:selectedMode" }
        }
    } # End else for normal script start

} catch { $errorMessage = "An unexpected error occurred: $($_.Exception.Message)"; Write-Error $errorMessage; Write-Host $errorMessage -ForegroundColor Red; try { [System.Windows.MessageBox]::Show($errorMessage, "Fatal Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) } catch {} }
finally {
    if ((Get-Variable -Name updateJob -Scope Script -EA SilentlyContinue) -ne $null -and $script:updateJob -ne $null) { Write-Host "`n--- Waiting for background update job (GPUpdate/CM Actions) to complete... ---" -ForegroundColor Yellow; Wait-Job $script:updateJob | Out-Null; Write-Host "--- Background Update Job Output (GPUpdate/CM Actions): ---" -ForegroundColor Yellow; Receive-Job $script:updateJob; Remove-Job $script:updateJob; Write-Host "--- End of Background Update Job Output ---" -ForegroundColor Yellow }
    else { Write-Host "`nNo background update job was started or it was already cleaned up." -ForegroundColor Gray }
}

Write-Host "--- Script Execution Finished ---"
if ($Host.Name -eq 'ConsoleHost' -and -not $psISE -and $env:TERM_PROGRAM -ne 'vscode') { if (-not $RunExpressElevated) { Write-Host "Press Enter to exit..." -ForegroundColor Yellow; Read-Host } }