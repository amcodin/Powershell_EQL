#requires -Version 3.0

# --- Load Assemblies FIRST ---
Write-Host "Attempting to load .NET Assemblies..."
try {
    Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop # Added for InputBox
    Write-Host ".NET Assemblies loaded successfully."
} catch {
    Write-Error "FATAL: Failed to load required .NET Assemblies. UI cannot be created."
    Write-Error "Error Message: $($_.Exception.Message)"
    # Optional: Pause if running directly in console
    if ($Host.Name -eq 'ConsoleHost' -and -not $psISE -and $env:TERM_PROGRAM -ne 'vscode') { Read-Host "Press Enter to exit" }
    Exit 1 # Exit script if assemblies fail to load
}

#region Global Variables & Configuration

# Define path templates using keys for easier management
# Use placeholders like {USERNAME} which will be replaced later
$script:pathTemplates = @{
    Signatures        = "{USERPROFILE}\AppData\Roaming\Microsoft\Signatures"
    # UserFolder        = "{USERPROFILE}" # Be cautious with this - very large! Commented out by default.
    QuickAccess       = "{USERPROFILE}\AppData\Roaming\Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms"
    # TempFolder        = "C:\Temp" # System-wide, not user-specific usually. Consider if needed.
    StickyNotesLegacy = "{USERPROFILE}\AppData\Roaming\Microsoft\Sticky Notes\StickyNotes.snt"
    StickyNotesModernDB = "{USERPROFILE}\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\plum.sqlite"
    GoogleEarthPlaces = "{USERPROFILE}\AppData\LocalLow\Google\GoogleEarth\myplaces.kml" # Corrected path to LocalLow
    ChromeBookmarks   = "{USERPROFILE}\AppData\Local\Google\Chrome\User Data\Default\Bookmarks"
    EdgeBookmarks     = "{USERPROFILE}\AppData\Local\Microsoft\Edge\User Data\Default\Bookmarks"
    # Add more templates here
}
$script:DefaultPath = "C:\LocalData" # Define default path globally

#endregion Global Variables & Configuration

#region Functions

# --- Helper Functions ---
function Format-Bytes {
    param([Parameter(Mandatory=$true)][long]$Bytes)
    $suffix = @("B", "KB", "MB", "GB", "TB", "PB"); $index = 0; $value = [double]$Bytes
    while ($value -ge 1024 -and $index -lt ($suffix.Length - 1)) { $value /= 1024; $index++ }
    return "{0:N2} {1}" -f $value, $suffix[$index]
}

# --- Path Resolution Functions ---

# Resolves path templates for a specific user profile path
function Resolve-UserPathTemplates {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserProfilePath, # e.g., C:\Users\jdoe or X:\Users\jdoe
        [Parameter(Mandatory=$true)]
        [hashtable]$Templates
    )

    $resolvedPaths = [System.Collections.Generic.List[PSCustomObject]]::new()
    $userName = Split-Path $UserProfilePath -Leaf # Extract username for logging

    Write-Host "Resolving paths based on profile: $UserProfilePath"

    foreach ($key in $Templates.Keys) {
        $template = $Templates[$key]
        # Replace placeholder with the actual profile path
        $resolvedPath = $template.Replace('{USERPROFILE}', $UserProfilePath.TrimEnd('\'))

        # Check if the resolved path exists
        if (Test-Path -LiteralPath $resolvedPath -ErrorAction SilentlyContinue) {
            $pathType = if (Test-Path -LiteralPath $resolvedPath -PathType Container) { "Folder" } else { "File" }
            Write-Host "  Found: '$key' -> $resolvedPath ($pathType)" -ForegroundColor Green
            $resolvedPaths.Add([PSCustomObject]@{
                Name = $key # Use the template key as the name
                Path = $resolvedPath
                Type = $pathType
                SourceKey = $key # Keep track of the original template key
            })
        } else {
            Write-Host "  Path not found or resolution failed: $resolvedPath (Derived from template key: $key)" -ForegroundColor Yellow
        }
    }
    return $resolvedPaths
}

# Gets standard user folders (Downloads, Desktop, etc.)
function Get-UserPaths {
    [CmdletBinding()]
    param (
        # No remote parameters needed for local backup context
    )
    $folderNames = @( "Downloads", "Pictures", "Videos" )
    $result = [System.Collections.Generic.List[PSCustomObject]]::new()
    $sourceBasePath = $env:USERPROFILE
    Write-Host "Getting local user folders from: $sourceBasePath"

    foreach ($folderName in $folderNames) {
        $fullPath = Join-Path -Path $sourceBasePath -ChildPath $folderName
        if (Test-Path -LiteralPath $fullPath -PathType Container) {
            $item = Get-Item -LiteralPath $fullPath -ErrorAction SilentlyContinue
             if ($item) {
                 $folderKey = "UserFolder_" + ($item.Name -replace '[^a-zA-Z0-9_]','_')
                 $result.Add([PSCustomObject]@{
                     Name = $folderKey # Use a generated key for Name/SourceKey
                     Path = $item.FullName; # Source Path (Local)
                     Type = "Folder";
                     SourceKey = $folderKey
                 })
             } else { Write-Host "User folder resolved but Get-Item failed: $fullPath" }
        } else { Write-Host "User folder not found or not a folder: $fullPath" }
    }
    return $result
}


# --- GUI Functions ---

# Show mode selection dialog (Backup/Restore/Express Only)
function Show-ModeDialog {
    Write-Host "Entering Show-ModeDialog function."
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Operation"
    $form.Size = New-Object System.Drawing.Size(400, 150) # Adjusted size
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.HelpButton = $false # Prevent closing via Esc without selection

    # No description label needed if buttons are clear

    $btnBackup = New-Object System.Windows.Forms.Button
    $btnBackup.Location = New-Object System.Drawing.Point(30, 40) # Adjusted position
    $btnBackup.Size = New-Object System.Drawing.Size(80, 30)
    $btnBackup.Text = "Backup"
    $btnBackup.DialogResult = [System.Windows.Forms.DialogResult]::Yes # Using Yes for Backup
    $form.Controls.Add($btnBackup)

    $btnRestore = New-Object System.Windows.Forms.Button
    $btnRestore.Location = New-Object System.Drawing.Point(140, 40) # Adjusted position
    $btnRestore.Size = New-Object System.Drawing.Size(80, 30)
    $btnRestore.Text = "Restore"
    $btnRestore.DialogResult = [System.Windows.Forms.DialogResult]::No # Using No for Restore
    $form.Controls.Add($btnRestore)

    $btnExpress = New-Object System.Windows.Forms.Button
    $btnExpress.Location = New-Object System.Drawing.Point(250, 40) # Adjusted position
    $btnExpress.Size = New-Object System.Drawing.Size(80, 30)
    $btnExpress.Text = "Express"
    $btnExpress.DialogResult = [System.Windows.Forms.DialogResult]::OK # Using OK for Express
    $form.Controls.Add($btnExpress)

    # No Cancel button

    # Set default button (optional, e.g., Express)
    $form.AcceptButton = $btnExpress
    # Allow closing via Esc key by setting CancelButton (maps to Restore in this case)
    # Or leave it unset if Esc should do nothing / close dialog with Cancel result
    # $form.CancelButton = $btnRestore # Example: Esc = Restore

    Write-Host "Showing mode selection dialog."
    $result = $form.ShowDialog()
    $form.Dispose()
    Write-Host "Mode selection dialog closed with result: $result"

    # Determine mode based on DialogResult
    $selectedMode = switch ($result) {
        ([System.Windows.Forms.DialogResult]::Yes) { 'Backup' }
        ([System.Windows.Forms.DialogResult]::No) { 'Restore' }
        ([System.Windows.Forms.DialogResult]::OK) { 'Express' }
        Default { 'Cancel' } # Includes Cancel or closing the dialog via 'X'
    }

    Write-Host "Determined mode: $selectedMode" -ForegroundColor Cyan

    # Handle Cancel explicitly (if user closes dialog with 'X')
    if ($selectedMode -eq 'Cancel') {
        Write-Host "Operation cancelled by user (dialog closed)." -ForegroundColor Yellow
        # Optional: Pause if running directly in console
        if ($Host.Name -eq 'ConsoleHost' -and -not $psISE -and $env:TERM_PROGRAM -ne 'vscode') { Read-Host "Press Enter to exit" }
        Exit 0 # Exit gracefully
    }

    # If restore mode, run updates immediately in a non-blocking way
    if ($selectedMode -eq 'Restore') {
        Write-Host "Restore mode selected. Initiating background system updates job..." -ForegroundColor Yellow
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
                 Write-Host "JOB: Attempting Method 1: Triggering via CIM ($clientSDKNamespace -> $clientClassName)..."
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
                 } catch { Write-Error "JOB: Unexpected error during CIM attempt: $($_.Exception.Message)"; $cimMethodSuccess = $false }
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
                 Write-Host "JOB: Exiting Start-ConfigManagerActions function."; return $overallSuccess
            }
            Set-GPupdate; Start-ConfigManagerActions; Write-Host "JOB: Background updates finished."
        }
        Write-Host "Background update job started (ID: $($script:updateJob.Id)). Output will be shown after main window closes. Main window will load now." -ForegroundColor Yellow
    }

    Write-Host "Exiting Show-ModeDialog function."
    return $selectedMode # Return the string 'Backup', 'Restore', or 'Express'
}

# Show main window (Backup/Restore specific)
function Show-MainWindow {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet('Backup', 'Restore')] # Only accepts Backup or Restore
        [string]$Mode
    )
    $IsBackup = ($Mode -eq 'Backup')
    $modeString = $Mode
    Write-Host "Entering Show-MainWindow function. Mode: $modeString"

    # XAML UI Definition (Added User Folders button and Space Labels)
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
        <!-- Added Space Labels -->
        <Label Name="lblFreeSpace" Content="Free Space: -" Margin="500,40,10,0" HorizontalAlignment="Right" VerticalAlignment="Top"/>
        <Label Name="lblRequiredSpace" Content="Required Space: -" Margin="500,65,10,0" HorizontalAlignment="Right" VerticalAlignment="Top"/>

        <Label Name="lblStatus" Content="Ready" Margin="10,0,10,10" HorizontalAlignment="Center" VerticalAlignment="Bottom" FontStyle="Italic"/>

        <Label Content="Files/Folders to Process:" Margin="10,90,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <ListView Name="lvwFiles" Margin="10,120,200,140" SelectionMode="Extended">
             <ListView.View>
                <GridView>
                    <!-- CheckBox Column -->
                    <GridViewColumn Width="30">
                        <GridViewColumn.CellTemplate>
                            <DataTemplate>
                                <CheckBox IsChecked="{Binding IsSelected, Mode=TwoWay}" />
                            </DataTemplate>
                        </GridViewColumn.CellTemplate>
                    </GridViewColumn>
                    <!-- Other Columns -->
                    <GridViewColumn Header="Name" DisplayMemberBinding="{Binding Name}" Width="180"/>
                    <GridViewColumn Header="Type" DisplayMemberBinding="{Binding Type}" Width="70"/>
                    <GridViewColumn Header="Path" DisplayMemberBinding="{Binding Path}" Width="280"/>
                </GridView>
            </ListView.View>
        </ListView>

        <StackPanel Margin="0,120,10,0" HorizontalAlignment="Right" Width="180">
            <Button Name="btnAddFile" Content="Add File" Width="120" Height="30" Margin="0,0,0,10"/>
            <Button Name="btnAddFolder" Content="Add Folder" Width="120" Height="30" Margin="0,0,0,10"/>
            <!-- Added User Folders Button -->
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

    try {
        Write-Host "Parsing XAML for main window."
        $reader = New-Object System.Xml.XmlNodeReader $XAML
        $window = [Windows.Markup.XamlReader]::Load($reader)
        Write-Host "XAML loaded successfully."

        Write-Host "Setting Window DataContext."
        $window.DataContext = [PSCustomObject]@{ IsRestoreMode = (-not $IsBackup) }

        Write-Host "Finding controls in main window."
        $controls = @{}
        @(
            'txtSaveLoc', 'btnBrowse', 'btnStart', 'lblMode', 'lblStatus',
            'lvwFiles', 'btnAddFile', 'btnAddFolder', 'btnRemove',
            'chkNetwork', 'chkPrinters', 'prgProgress', 'txtProgress',
            'btnAddBAUPaths', 'lblFreeSpace', 'lblRequiredSpace' # Added new controls
        ) | ForEach-Object { $controls[$_] = $window.FindName($_) }
        Write-Host "Controls found and stored in hashtable."

        # --- Space Calculation Script Blocks (defined in function scope) ---
        $UpdateFreeSpaceLabel = {
            param($ControlsParam)
            $location = $ControlsParam.txtSaveLoc.Text
            $freeSpaceString = "Free Space: N/A"
            if (-not [string]::IsNullOrEmpty($location)) {
                try {
                    $driveLetter = $null
                    if ($location -match '^[a-zA-Z]:\\') { $driveLetter = $location.Substring(0, 2) }
                    elseif ($location -match '^\\\\[^\\]+\\[^\\]+') { $freeSpaceString = "Free Space: N/A (UNC)" }

                    if ($driveLetter) {
                        $driveInfo = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$driveLetter'" -ErrorAction SilentlyContinue
                        if ($driveInfo -and $driveInfo.FreeSpace -ne $null) {
                            $freeSpaceString = "Free Space: $(Format-Bytes $driveInfo.FreeSpace)"
                        } else { Write-Warning "Could not get free space for drive $driveLetter" }
                    }
                } catch { Write-Warning "Error getting free space for '$location': $($_.Exception.Message)" }
            }
            # Update UI on Dispatcher thread
            $ControlsParam.lblFreeSpace.Dispatcher.InvokeAsync({ $ControlsParam.lblFreeSpace.Content = $freeSpaceString }) | Out-Null
            Write-Host "Updated Free Space Label: $freeSpaceString"
        }

        $UpdateRequiredSpaceLabel = {
            param($ControlsParam)
            $totalSize = 0L
            $requiredSpaceString = "Required Space: Calculating..."
            # Update UI immediately to show "Calculating..."
            $ControlsParam.lblRequiredSpace.Dispatcher.InvokeAsync({ $ControlsParam.lblRequiredSpace.Content = $requiredSpaceString }) | Out-Null

            # Perform calculation (can take time)
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
                            } else {
                                $fileSize = (Get-Item -LiteralPath $item.Path -Force -ErrorAction SilentlyContinue).Length
                                if ($fileSize -ne $null) { $totalSize += $fileSize }
                            }
                        } catch { Write-Warning "Error calculating size for '$($item.Path)': $($_.Exception.Message)" }
                    } else { Write-Warning "Checked item path not found, skipping size calculation: $($item.Path)" }
                }
                $requiredSpaceString = "Required Space: $(Format-Bytes $totalSize)"
            } else {
                $requiredSpaceString = "Required Space: 0 B"
            }
            # Update UI with final result on Dispatcher thread
            $ControlsParam.lblRequiredSpace.Dispatcher.InvokeAsync({ $ControlsParam.lblRequiredSpace.Content = $requiredSpaceString }) | Out-Null
            Write-Host "Updated Required Space Label: $requiredSpaceString"
        }
        # --- End Space Calculation Script Blocks ---


        # --- Window Initialization ---
        Write-Host "Initializing window controls based on mode."
        $controls.lblMode.Content = "Mode: $modeString"
        $controls.btnStart.Content = $modeString # "Backup" or "Restore"
        $controls.btnAddBAUPaths.IsEnabled = $IsBackup # Only enable "Add User Folders" in Backup mode

        # Set default path
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
        } else { Write-Host "Default path exists."}
        $controls.txtSaveLoc.Text = $script:DefaultPath
        # Update free space label initially
        & $UpdateFreeSpaceLabel -ControlsParam $controls


        # --- Load initial items based on mode ---
        Write-Host "Loading initial items for ListView based on mode."
        $itemsList = [System.Collections.Generic.List[PSCustomObject]]::new()

        if ($IsBackup) {
            Write-Host "Backup Mode: Getting default paths using local user profile."
            $localUserProfile = $env:USERPROFILE
            # Use Resolve-UserPathTemplates with local profile
            $paths = Resolve-UserPathTemplates -UserProfilePath $localUserProfile -Templates $script:pathTemplates

            if ($paths -ne $null -and $paths.Count -gt 0) {
                 Write-Host "Found $($paths.Count) local items. Adding 'IsSelected' property..."
                 $paths | ForEach-Object {
                     $_ | Add-Member -MemberType NoteProperty -Name 'IsSelected' -Value $true -PassThru
                     $itemsList.Add($_)
                 }
                 Write-Host "Populated ListView with local items."
            } else {
                 Write-Warning "Could not resolve any default local paths for backup."
            }
        }
        elseif (Test-Path $script:DefaultPath) { # Restore Mode
            Write-Host "Restore Mode: Checking for latest backup in $script:DefaultPath."
            $latestBackup = Get-ChildItem -Path $script:DefaultPath -Directory -Filter "Backup_*" |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1

            if ($latestBackup) {
                Write-Host "Found latest backup: $($latestBackup.FullName)"
                $controls.txtSaveLoc.Text = $latestBackup.FullName
                & $UpdateFreeSpaceLabel -ControlsParam $controls # Update free space for backup location
                $logFilePath = Join-Path $latestBackup.FullName "FileList_Backup.csv"
                if (Test-Path $logFilePath) {
                    # Load backup contents into ListView
                    $backupItems = Get-ChildItem -Path $latestBackup.FullName |
                        Where-Object { $_.PSIsContainer -or $_.Name -match '^(Drives\.csv|Printers\.txt)$' } | # List folders or known settings files
                        Where-Object { $_.Name -notmatch '^(FileList_.*\.csv|TransferLog\.csv)$' } | # Exclude logs
                        ForEach-Object {
                            [PSCustomObject]@{
                                Name = $_.Name # Name represents the SourceKey or setting type
                                Type = if ($_.PSIsContainer) { "Folder" } else { "Setting" }
                                Path = $_.FullName # Path within the backup folder
                                IsSelected = $true # Default to selected for restore
                            }
                        }
                    $backupItems | ForEach-Object { $itemsList.Add($_) }
                    Write-Host "Populated ListView with $($itemsList.Count) items/categories from latest backup."
                } else {
                    Write-Warning "Latest backup folder '$($latestBackup.FullName)' is missing log file '$logFilePath'."
                    $controls.lblStatus.Content = "Restore mode: Latest backup folder is invalid. Please browse."
                }
            } else {
                 Write-Host "No backups found in $script:DefaultPath."
                 $controls.lblStatus.Content = "Restore mode: No backups found in $script:DefaultPath. Please browse."
            }
        } else { # Restore Mode, default path doesn't exist
             Write-Host "Restore Mode: Default path $script:DefaultPath does not exist."
             $controls.lblStatus.Content = "Restore mode: Default path $script:DefaultPath does not exist. Please browse."
        }

        # Assign items to ListView
        if ($controls['lvwFiles'] -ne $null) {
            $controls.lvwFiles.ItemsSource = $itemsList
            Write-Host "Assigned $($itemsList.Count) items to ListView ItemsSource."
            # Update required space after loading items
            & $UpdateRequiredSpaceLabel -ControlsParam $controls
        } else {
            Write-Error "ListView control ('lvwFiles') not found!"
            throw "ListView control ('lvwFiles') could not be found in the XAML."
        }
        Write-Host "Finished loading initial items."


        # --- Event Handlers ---
        Write-Host "Assigning event handlers."
        $controls.btnBrowse.Add_Click({
            Write-Host "Browse button clicked."
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $dialog.Description = if ($IsBackup) { "Select location to save backup" } else { "Select backup folder to restore from" }
            if(Test-Path $controls.txtSaveLoc.Text){
                 $dialog.SelectedPath = $controls.txtSaveLoc.Text
            } else {
                 $dialog.SelectedPath = $script:DefaultPath # Fallback if current text isn't valid
            }
            $dialog.ShowNewFolderButton = $IsBackup # Only allow creating new folders in backup mode

            $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized' }
            Write-Host "Showing FolderBrowserDialog."
            $result = $dialog.ShowDialog($owner)
            $owner.Dispose() # Dispose the temporary owner form

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $selectedPath = $dialog.SelectedPath
                Write-Host "Folder selected: $selectedPath"
                $controls.txtSaveLoc.Text = $selectedPath
                & $UpdateFreeSpaceLabel -ControlsParam $controls # Update free space for new location

                if (-not $IsBackup) { # Restore Mode - Reload list from selected backup
                    Write-Host "Restore Mode: Loading items from selected backup folder."
                    $logFilePath = Join-Path $selectedPath "FileList_Backup.csv"
                    Write-Host "Checking for log file: $logFilePath"
                    $newItemsList = [System.Collections.Generic.List[PSCustomObject]]::new() # Clear list first
                    if (Test-Path -Path $logFilePath) {
                         Write-Host "Log file found. Populating ListView."
                         $backupItems = Get-ChildItem -Path $selectedPath |
                            Where-Object { $_.PSIsContainer -or $_.Name -match '^(Drives\.csv|Printers\.txt)$' } |
                            Where-Object { $_.Name -notmatch '^(FileList_.*\.csv|TransferLog\.csv)$' } |
                            ForEach-Object {
                                [PSCustomObject]@{ Name = $_.Name; Type = if ($_.PSIsContainer) { "Folder" } else { "Setting" }; Path = $_.FullName; IsSelected = $true }
                            }
                        $backupItems | ForEach-Object { $newItemsList.Add($_) }
                        $controls.lblStatus.Content = "Ready to restore from: $selectedPath"
                        Write-Host "ListView updated with $($newItemsList.Count) items/categories from selected backup."
                    } else {
                         Write-Warning "Selected folder is not a valid backup (missing FileList_Backup.csv)."
                         $controls.lblStatus.Content = "Selected folder is not a valid backup (missing FileList_Backup.csv)."
                         [System.Windows.MessageBox]::Show("The selected folder does not appear to be a valid backup. It's missing the 'FileList_Backup.csv' log file.", "Invalid Backup Folder", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    }
                    $controls.lvwFiles.ItemsSource = $newItemsList # Assign new list (or empty list if invalid)
                    & $UpdateRequiredSpaceLabel -ControlsParam $controls # Update required space
                } else { # Backup Mode
                     $controls.lblStatus.Content = "Backup location set to: $selectedPath"
                     # Required space doesn't change just by changing destination
                }
            } else { Write-Host "Folder selection cancelled."}
        })

        # Add File handler
        $controls.btnAddFile.Add_Click({
            if (!$IsBackup) { [System.Windows.MessageBox]::Show("Adding individual files is only supported in Backup mode.", "Information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information); return }
            Write-Host "Add File button clicked."
            $dialog = New-Object System.Windows.Forms.OpenFileDialog; $dialog.Title = "Select File(s) to Add for Backup"; $dialog.Multiselect = $true
            $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized' }; $result = $dialog.ShowDialog($owner); $owner.Dispose()
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                Write-Host "$($dialog.FileNames.Count) file(s) selected."
                $currentItems = $controls.lvwFiles.ItemsSource -as [System.Collections.Generic.List[PSCustomObject]]
                if ($currentItems -eq $null) { $currentItems = [System.Collections.Generic.List[PSCustomObject]]::new(); $controls.lvwFiles.ItemsSource = $currentItems }
                $addedCount = 0
                foreach ($file in $dialog.FileNames) {
                    if (-not ($currentItems.Path -contains $file)) {
                        $fileKey = "ManualFile_" + ([System.IO.Path]::GetFileNameWithoutExtension($file) -replace '[^a-zA-Z0-9_]','_'); Write-Host "Adding file: $file (Key: $fileKey)"
                        $currentItems.Add([PSCustomObject]@{ Name = $fileKey; Type = "File"; Path = $file; IsSelected = $true; SourceKey = $fileKey }); $addedCount++
                    } else { Write-Host "Skipping duplicate file: $file"}
                }
                if ($addedCount -gt 0) { $controls.lvwFiles.Items.Refresh(); & $UpdateRequiredSpaceLabel -ControlsParam $controls } # Refresh view and update space
                Write-Host "Updated ListView ItemsSource. Added $addedCount new file(s)."
            } else { Write-Host "File selection cancelled."}
        })

        # Add Folder handler
        $controls.btnAddFolder.Add_Click({
            if (!$IsBackup) { [System.Windows.MessageBox]::Show("Adding individual folders is only supported in Backup mode.", "Information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information); return }
            Write-Host "Add Folder button clicked."
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog; $dialog.Description = "Select Folder to Add for Backup"; $dialog.ShowNewFolderButton = $false
            $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized' }; $result = $dialog.ShowDialog($owner); $owner.Dispose()
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                 $selectedPath = $dialog.SelectedPath; Write-Host "Folder selected to add: $selectedPath"
                 $currentItems = $controls.lvwFiles.ItemsSource -as [System.Collections.Generic.List[PSCustomObject]]
                 if ($currentItems -eq $null) { $currentItems = [System.Collections.Generic.List[PSCustomObject]]::new(); $controls.lvwFiles.ItemsSource = $currentItems }
                 if (-not ($currentItems.Path -contains $selectedPath)) {
                    $folderKey = "ManualFolder_" + ([System.IO.Path]::GetFileName($selectedPath) -replace '[^a-zA-Z0-9_]','_'); Write-Host "Adding folder: $selectedPath (Key: $folderKey)"
                    $currentItems.Add([PSCustomObject]@{ Name = $folderKey; Type = "Folder"; Path = $selectedPath; IsSelected = $true; SourceKey = $folderKey })
                    $controls.lvwFiles.Items.Refresh(); & $UpdateRequiredSpaceLabel -ControlsParam $controls # Refresh view and update space
                    Write-Host "Updated ListView ItemsSource with new folder."
                 } else { Write-Host "Skipping duplicate folder: $selectedPath" }
            } else { Write-Host "Folder selection cancelled."}
        })

        # Add User Folders (BAU Paths) handler
        $controls.btnAddBAUPaths.Add_Click({
            if (!$IsBackup) { return } # Should be disabled anyway, but double-check
            Write-Host "Add User Folders button clicked."
            try {
                $userPaths = Get-UserPaths # Gets local user folders
                if ($userPaths -and $userPaths.Count -gt 0) {
                    $currentItems = $controls.lvwFiles.ItemsSource -as [System.Collections.Generic.List[PSCustomObject]]
                    if ($currentItems -eq $null) { $currentItems = [System.Collections.Generic.List[PSCustomObject]]::new(); $controls.lvwFiles.ItemsSource = $currentItems }
                    $addedCount = 0
                    foreach ($pathInfo in $userPaths) {
                        if (-not ($currentItems.Path -contains $pathInfo.Path)) {
                            Write-Host "Adding User Folder: $($pathInfo.Path) (Key: $($pathInfo.SourceKey))"
                            $pathInfo | Add-Member -MemberType NoteProperty -Name 'IsSelected' -Value $true -PassThru
                            $currentItems.Add($pathInfo)
                            $addedCount++
                        } else { Write-Host "Skipping duplicate user folder: $($pathInfo.Path)" }
                    }
                    if ($addedCount -gt 0) { $controls.lvwFiles.Items.Refresh(); & $UpdateRequiredSpaceLabel -ControlsParam $controls } # Refresh view and update space
                    Write-Host "Added $addedCount new user folder(s)."
                } else { Write-Warning "Could not find any standard user folders." }
            } catch { Write-Error "Error adding user folders: $($_.Exception.Message)" }
        })

        # Remove Selected handler
        $controls.btnRemove.Add_Click({
            Write-Host "Remove Selected button clicked."
            $selectedObjects = @($controls.lvwFiles.SelectedItems)
            if ($selectedObjects.Count -gt 0) {
                Write-Host "Removing $($selectedObjects.Count) selected item(s)."
                $currentItems = $controls.lvwFiles.ItemsSource -as [System.Collections.Generic.List[PSCustomObject]]
                if ($currentItems -ne $null) {
                    $itemsToRemove = $selectedObjects | ForEach-Object { $_ }
                    $itemsToRemove | ForEach-Object { $currentItems.Remove($_) } | Out-Null
                    $controls.lvwFiles.Items.Refresh(); & $UpdateRequiredSpaceLabel -ControlsParam $controls # Refresh view and update space
                    Write-Host "ListView ItemsSource updated after removal."
                }
            } else { Write-Host "No items selected to remove."}
        })

        # Event handler for ListView selection change (to update required space)
        $controls.lvwFiles.Add_SelectionChanged({
             Write-Host "ListView Selection Changed - Updating Required Space"
             # Run calculation in background to avoid UI lag if list is huge
             [System.Threading.Tasks.Task]::Run( { & $UpdateRequiredSpaceLabel -ControlsParam $controls } ) | Out-Null
        })
        # Also update space when checkbox is clicked (SelectionChanged doesn't always fire reliably for checkbox clicks)
        $controls.lvwFiles.Add_PreviewMouseLeftButtonUp({
            param($sender, $e)
            # Check if the click was on a CheckBox
            $originalSource = $e.OriginalSource
            if ($originalSource -is [System.Windows.Controls.CheckBox]) {
                Write-Host "ListView CheckBox Clicked - Updating Required Space"
                # Use Dispatcher.BeginInvoke to allow the checkbox state to update *before* calculating
                $window.Dispatcher.BeginInvoke([action]{
                    [System.Threading.Tasks.Task]::Run( { & $UpdateRequiredSpaceLabel -ControlsParam $controls } ) | Out-Null
                }, [System.Windows.Threading.DispatcherPriority]::Background)
            }
        })


        # --- Start Button Logic (Backup/Restore) ---
        $controls.btnStart.Add_Click({
            $modeStringLocal = if ($IsBackup) { 'Backup' } else { 'Restore' } # Use local var to avoid scope issues in event handler
            Write-Host "Start button clicked. Mode: $modeStringLocal"

            $location = $controls.txtSaveLoc.Text
            Write-Host "Selected location: $location"
            if ([string]::IsNullOrEmpty($location) -or -not (Test-Path $location -PathType Container)) {
                Write-Warning "Invalid location selected."
                [System.Windows.MessageBox]::Show("Please select a valid target directory first.", "Location Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }

            # Final check/update of required space before starting
            Write-Host "Final check of required space..."
            & $UpdateRequiredSpaceLabel -ControlsParam $controls

            # Get items to process based on selection
            $itemsToProcess = [System.Collections.Generic.List[PSCustomObject]]::new()
            $checkedItems = @($controls.lvwFiles.ItemsSource) | Where-Object { $_.IsSelected }
            if ($checkedItems) { $checkedItems | ForEach-Object { $itemsToProcess.Add($_) } }

            $doNetwork = $controls.chkNetwork.IsChecked
            $doPrinters = $controls.chkPrinters.IsChecked

            Write-Host "Disabling UI controls and setting wait cursor."
            $controls | ForEach-Object { if ($_.Value -is [System.Windows.Controls.Control]) { $_.Value.IsEnabled = $false } }
            $window.Cursor = [System.Windows.Input.Cursors]::Wait

            # Define UI Progress Action ScriptBlock
            $uiProgressAction = {
                param($status, $percent, $details)
                # Ensure UI updates happen on the Dispatcher thread
                $window.Dispatcher.InvokeAsync( [action]{
                    $controls.lblStatus.Content = $status
                    $controls.txtProgress.Text = $details
                    if ($percent -ge 0) { # Normal progress
                        $controls.prgProgress.IsIndeterminate = $false
                        $controls.prgProgress.Value = $percent
                    } elseif ($percent -eq -1) { # Error state
                        $controls.prgProgress.IsIndeterminate = $false
                        $controls.prgProgress.Value = 0
                    } else { # Indeterminate state (e.g., percent = -2)
                        $controls.prgProgress.IsIndeterminate = $true
                    }
                }, [System.Windows.Threading.DispatcherPriority]::Background ) | Out-Null
            }

            # Start Operation in a separate task/thread to keep UI responsive
            $operationTask = [System.Threading.Tasks.Task]::Factory.StartNew({
                $success = $false
                try {
                    & $uiProgressAction.Invoke("Initializing...", -2, "Starting operation...") # Indeterminate progress

                    if ($IsBackup) {
                        Write-Host "--- Starting Background Backup Operation ---"
                        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                        $localUsername = $env:USERNAME -replace '[^a-zA-Z0-9]', '_'
                        $backupRootPath = Join-Path $location "Backup_${localUsername}_$timestamp"
                        if ($itemsToProcess.Count -eq 0 -and !$doNetwork -and !$doPrinters) { throw "No items or settings selected for backup." }

                        # Call Backup Function (Refactored for background)
                        $success = Invoke-BackupOperation -BackupRootPath $backupRootPath -ItemsToBackup $itemsToProcess -BackupNetworkDrives $doNetwork -BackupPrinters $doPrinters -ProgressAction $uiProgressAction

                    } else { # Restore Mode
                        Write-Host "--- Starting Background Restore Operation ---"
                        $backupRootPath = $location
                        # Get selected keys/files for filtering restore
                        $selectedKeysOrFilesForRestore = $itemsToProcess | Select-Object -ExpandProperty Name

                        # Call Restore Function (Refactored for background)
                        $success = Invoke-RestoreOperation -BackupRootPath $backupRootPath -SelectedKeysOrFiles $selectedKeysOrFilesForRestore -RestoreNetworkDrives $doNetwork -RestorePrinters $doPrinters -ProgressAction $uiProgressAction
                    }

                    # --- Operation Completion (Background Thread) ---
                    if ($success) {
                        Write-Host "Background operation completed successfully." -ForegroundColor Green
                        & $uiProgressAction.Invoke("Operation Completed", 100, "Successfully finished $modeStringLocal.")
                        # Show message box from UI thread
                        $window.Dispatcher.InvokeAsync({ [System.Windows.MessageBox]::Show("The $modeStringLocal operation completed successfully!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) }) | Out-Null
                    } else {
                        Write-Error "Background operation failed."
                        & $uiProgressAction.Invoke("Operation Failed", -1, "Errors occurred during $modeStringLocal. Check console.")
                        $window.Dispatcher.InvokeAsync({ [System.Windows.MessageBox]::Show("The $modeStringLocal operation failed. Check console output for details.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) }) | Out-Null
                    }

                } catch {
                    # --- Operation Failure (Background Thread) ---
                    $errorMessage = "Operation Failed (Background Task): $($_.Exception.Message)"
                    Write-Error $errorMessage
                    & $uiProgressAction.Invoke("Operation Failed", -1, $errorMessage)
                    $window.Dispatcher.InvokeAsync({ [System.Windows.MessageBox]::Show($errorMessage, "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) }) | Out-Null
                } finally {
                    # --- Re-enable UI (Background Thread -> Dispatcher) ---
                    Write-Host "Operation finished (background finally block). Re-enabling UI controls via Dispatcher."
                    $window.Dispatcher.InvokeAsync({
                        $controls | ForEach-Object { if ($_.Value -is [System.Windows.Controls.Control]) {
                                # Special handling for Add User Folders button
                                if ($_.Key -eq 'btnAddBAUPaths') { $_.Value.IsEnabled = $IsBackup }
                                else { $_.Value.IsEnabled = $true }
                            }
                        }
                        $window.Cursor = [System.Windows.Input.Cursors]::Arrow
                        Write-Host "UI Controls re-enabled and cursor reset."
                    }) | Out-Null
                }
            }) # End Task Factory StartNew

            # Optional: Handle task exceptions if needed, though the inner try/catch should handle most
            $operationTask.ContinueWith({
                param($task)
                if ($task.IsFaulted) {
                    $aggEx = $task.Exception.Flatten()
                    $aggEx.InnerExceptions | ForEach-Object { Write-Error "Unhandled Task Exception: $($_.Message)" }
                }
            }, [System.Threading.Tasks.TaskContinuationOptions]::OnlyOnFaulted)

        }) # End btnStart.Add_Click


        # --- Show Window ---
        Write-Host "Showing main window."
        $window.ShowDialog() | Out-Null
        Write-Host "Main window closed."

    } catch {
        # --- Window Load Failure ---
        $errorMessage = "Failed to load main window: $($_.Exception.Message)"
        Write-Error $errorMessage
        Write-Host "FATAL ERROR: Failed to load main window: $errorMessage" -ForegroundColor Red
        try {
             [System.Windows.MessageBox]::Show($errorMessage, "Critical Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        } catch {
             Write-Warning "Could not display error message box."
        }
    } finally {
         Write-Host "Exiting Show-MainWindow function."
         # Clean up script-scoped variables if they exist
         Remove-Variable -Name UpdateFreeSpaceLabel -Scope Script -ErrorAction SilentlyContinue
         Remove-Variable -Name UpdateRequiredSpaceLabel -Scope Script -ErrorAction SilentlyContinue
    }
}

# --- Refactored Backup/Restore Functions (Called by UI or Express) ---

function Invoke-BackupOperation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)] [string]$BackupRootPath,
        [Parameter(Mandatory=$true)] [System.Collections.Generic.List[PSCustomObject]]$ItemsToBackup,
        [Parameter(Mandatory=$true)] [bool]$BackupNetworkDrives,
        [Parameter(Mandatory=$true)] [bool]$BackupPrinters,
        [Parameter(Mandatory=$false)] [scriptblock]$ProgressAction # Optional: $ProgressAction.Invoke($status, $percent, $details)
    )
    $UpdateProgress = { if ($ProgressAction) { $ProgressAction.Invoke($args[0], $args[1], $args[2]) } else { Write-Host "$($args[0]) ($($args[1])%) - $($args[2])" } }
    Write-Host "--- Starting LOCAL Backup Operation to '$BackupRootPath' ---"
    try {
        $UpdateProgress.Invoke("Initializing Backup", 0, "Creating backup directory...")
        if (-not (Test-Path $BackupRootPath)) { New-Item -Path $BackupRootPath -ItemType Directory -Force -EA Stop | Out-Null; Write-Host "Backup directory created." }
        $csvLogPath = Join-Path $BackupRootPath "FileList_Backup.csv"; "OriginalFullPath,BackupRelativePath,SourceKey" | Set-Content -Path $csvLogPath -Encoding UTF8
        if (-not $ItemsToBackup -or $ItemsToBackup.Count -eq 0) { throw "No items specified for backup." }
        $networkDriveCount = $(if ($BackupNetworkDrives) { 1 } else { 0 }); $printerCount = $(if ($BackupPrinters) { 1 } else { 0 })
        $totalItems = $ItemsToBackup.Count + $networkDriveCount + $printerCount; $currentItemIndex = 0
        foreach ($item in $ItemsToBackup) {
            $currentItemIndex++; $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 90) } else { 0 } # File copy up to 90%
            $statusMessage = "Backing up Item $currentItemIndex of $totalItems"; $UpdateProgress.Invoke($statusMessage, $percentComplete, "Processing: $($item.Name)")
            $sourcePath = $item.Path; $sourceKey = $item.SourceKey
            if (-not (Test-Path -LiteralPath $sourcePath)) { Write-Warning "SKIP: $sourcePath not found."; "`"$sourcePath`",`"SKIPPED_NOT_FOUND`",`"$sourceKey`"" | Add-Content -Path $csvLogPath -Encoding UTF8; continue }
            $backupItemRootName = $sourceKey -replace '[^a-zA-Z0-9_]', '_'
            try {
                if ($item.Type -eq "Folder") {
                    $sourceFolderInfo = Get-Item -LiteralPath $sourcePath
                    Get-ChildItem -LiteralPath $sourcePath -Recurse -File -Force -EA Stop | ForEach-Object {
                        $originalFileFullPath = $_.FullName; $relativeFilePath = $originalFileFullPath.Substring($sourceFolderInfo.FullName.Length).TrimStart('\'); $backupRelativePath = Join-Path $backupItemRootName $relativeFilePath
                        $targetBackupPath = Join-Path $BackupRootPath $backupRelativePath; $targetBackupDir = Split-Path $targetBackupPath -Parent
                        if (-not (Test-Path $targetBackupDir)) { New-Item -Path $targetBackupDir -ItemType Directory -Force -EA Stop | Out-Null }
                        Copy-Item -LiteralPath $originalFileFullPath -Destination $targetBackupPath -Force -EA Stop
                        "`"$originalFileFullPath`",`"$backupRelativePath`",`"$sourceKey`"" | Add-Content -Path $csvLogPath -Encoding UTF8
                        $UpdateProgress.Invoke($statusMessage, $percentComplete, "Copied: $($_.Name)")
                    }
                } else { # File
                    $originalFileFullPath = $sourcePath; $backupRelativePath = Join-Path $backupItemRootName (Split-Path $originalFileFullPath -Leaf); $targetBackupPath = Join-Path $BackupRootPath $backupRelativePath; $targetBackupDir = Split-Path $targetBackupPath -Parent
                    if (-not (Test-Path $targetBackupDir)) { New-Item -Path $targetBackupDir -ItemType Directory -Force -EA Stop | Out-Null }
                    Copy-Item -LiteralPath $originalFileFullPath -Destination $targetBackupPath -Force -EA Stop
                    "`"$originalFileFullPath`",`"$backupRelativePath`",`"$sourceKey`"" | Add-Content -Path $csvLogPath -Encoding UTF8
                }
            } catch { $errMsg = $_.Exception.Message -replace '"','""'; Write-Warning "ERROR copying '$($item.Name)': $errMsg"; "`"$sourcePath`",`"ERROR_COPY: $errMsg`",`"$sourceKey`"" | Add-Content -Path $csvLogPath -Encoding UTF8 }
        }
        if ($BackupNetworkDrives) {
            $currentItemIndex++; $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 95) } else { 0 }; $UpdateProgress.Invoke("Backing up Settings", $percentComplete, "Network Drives...")
            try { Get-WmiObject -Class Win32_MappedLogicalDisk -EA Stop | Select-Object Name, ProviderName | Export-Csv -Path (Join-Path $BackupRootPath "Drives.csv") -NoTypeInformation -Encoding UTF8 -EA Stop; Write-Host "Drives backed up." }
            catch { Write-Warning "Failed to backup drives: $($_.Exception.Message)" }
        }
        if ($BackupPrinters) {
            $currentItemIndex++; $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 100) } else { 0 }; $UpdateProgress.Invoke("Backing up Settings", $percentComplete, "Printers...")
            try { Get-WmiObject -Class Win32_Printer -Filter "Local = False AND Network = True" -EA Stop | Select-Object -ExpandProperty Name | Set-Content -Path (Join-Path $BackupRootPath "Printers.txt") -Encoding UTF8 -EA Stop; Write-Host "Printers backed up." }
            catch { Write-Warning "Failed to backup printers: $($_.Exception.Message)" }
        }
        $UpdateProgress.Invoke("Backup Complete", 100, "Successfully backed up to: $BackupRootPath"); Write-Host "--- LOCAL Backup Operation Finished ---"; return $true
    } catch { $errorMessage = "LOCAL Backup Failed: $($_.Exception.Message)"; Write-Error $errorMessage; $UpdateProgress.Invoke("Backup Failed", -1, $errorMessage); return $false }
}

function Invoke-RestoreOperation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)] [string]$BackupRootPath,
        [Parameter(Mandatory=$true)] [string[]]$SelectedKeysOrFiles, # Names from the ListView selection
        [Parameter(Mandatory=$true)] [bool]$RestoreNetworkDrives,
        [Parameter(Mandatory=$true)] [bool]$RestorePrinters,
        [Parameter(Mandatory=$false)] [scriptblock]$ProgressAction
    )
    $UpdateProgress = { if ($ProgressAction) { $ProgressAction.Invoke($args[0], $args[1], $args[2]) } else { Write-Host "$($args[0]) ($($args[1])%) - $($args[2])" } }
    Write-Host "--- Starting LOCAL Restore Operation from '$BackupRootPath' ---"
    try {
        $UpdateProgress.Invoke("Initializing Restore", 0, "Checking backup contents...")
        $csvLogPath = Join-Path $BackupRootPath "FileList_Backup.csv"; $drivesCsvPath = Join-Path $BackupRootPath "Drives.csv"; $printersTxtPath = Join-Path $BackupRootPath "Printers.txt"
        if (-not (Test-Path $csvLogPath -PathType Leaf)) { throw "Backup log file 'FileList_Backup.csv' not found." }
        $backupLog = Import-Csv -Path $csvLogPath -Encoding UTF8; if (-not $backupLog) { throw "Backup log is empty." }
        if (-not ($backupLog | Get-Member -Name SourceKey)) { throw "Backup log missing 'SourceKey' column." }
        $logEntriesToRestore = $backupLog | Where-Object { $_.SourceKey -in $SelectedKeysOrFiles -and $_.BackupRelativePath -notmatch '^(SKIPPED|ERROR)_' }
        if (-not $logEntriesToRestore) { Write-Warning "No valid file/folder entries in log match selection." }
        $networkDriveCount = $(if ($RestoreNetworkDrives -and ($SelectedKeysOrFiles -contains "Drives.csv")) { 1 } else { 0 })
        $printerCount = $(if ($RestorePrinters -and ($SelectedKeysOrFiles -contains "Printers.txt")) { 1 } else { 0 })
        $totalItems = $logEntriesToRestore.Count + $networkDriveCount + $printerCount; $currentItemIndex = 0
        if ($logEntriesToRestore.Count -gt 0) {
            Write-Host "Restoring $($logEntriesToRestore.Count) file/folder entries..."
            foreach ($entry in $logEntriesToRestore) {
                $currentItemIndex++; $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 90) } else { 0 } # Restore up to 90%
                $statusMessage = "Restoring Item $currentItemIndex of $totalItems"; $originalFileFullPath = $entry.OriginalFullPath; $sourceBackupPath = Join-Path $BackupRootPath $entry.BackupRelativePath
                $UpdateProgress.Invoke($statusMessage, $percentComplete, "Restoring: $(Split-Path $originalFileFullPath -Leaf)")
                if (Test-Path -LiteralPath $sourceBackupPath -PathType Leaf) {
                    try { $targetRestoreDir = Split-Path $originalFileFullPath -Parent; if (-not (Test-Path $targetRestoreDir)) { New-Item -Path $targetRestoreDir -ItemType Directory -Force -EA Stop | Out-Null }; Copy-Item -LiteralPath $sourceBackupPath -Destination $originalFileFullPath -Force -EA Stop }
                    catch { Write-Warning "Failed to restore '$originalFileFullPath': $($_.Exception.Message)" }
                } else { Write-Warning "SKIP: Source '$sourceBackupPath' not found." }
            }
        }
        if ($RestoreNetworkDrives -and ($SelectedKeysOrFiles -contains "Drives.csv")) {
            $currentItemIndex++; $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 95) } else { 0 }; $UpdateProgress.Invoke("Restoring Settings", $percentComplete, "Network Drives...")
            if (Test-Path $drivesCsvPath) {
                try { Import-Csv $drivesCsvPath | ForEach-Object { $driveLetter = $_.Name.TrimEnd(':'); $networkPath = $_.ProviderName; if ($driveLetter -match '^[A-Z]$' -and $networkPath -match '^\\\\') { if (-not (Test-Path -LiteralPath "$($driveLetter):")) { try { Write-Host "  Mapping $driveLetter -> $networkPath"; New-PSDrive -Name $driveLetter -PSProvider FileSystem -Root $networkPath -Persist -Scope Global -EA Stop } catch { Write-Warning "  Failed map $driveLetter`: $($_.Exception.Message)" } } else { Write-Host "  Skip $driveLetter`: exists." } } else { Write-Warning "  Skip invalid map: $($_.Name)" } } }
                catch { Write-Warning "Error restoring drives: $($_.Exception.Message)" }
            } else { Write-Warning "Drives.csv selected but not found." }
        }
        if ($RestorePrinters -and ($SelectedKeysOrFiles -contains "Printers.txt")) {
            $currentItemIndex++; $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 100) } else { 0 }; $UpdateProgress.Invoke("Restoring Settings", $percentComplete, "Printers...")
            if (Test-Path $printersTxtPath) {
                try { $wsNet = New-Object -ComObject WScript.Network; Get-Content $printersTxtPath | ForEach-Object { $printerPath = $_.Trim(); if (-not ([string]::IsNullOrWhiteSpace($printerPath)) -and $printerPath -match '^\\\\') { try { Write-Host "  Adding printer: $printerPath"; $wsNet.AddWindowsPrinterConnection($printerPath) } catch { Write-Warning "  Failed add printer '$printerPath': $($_.Exception.Message)" } } else { Write-Warning "  Skip invalid printer line: '$_'" } } }
                catch { Write-Warning "Error restoring printers: $($_.Exception.Message)" }
            } else { Write-Warning "Printers.txt selected but not found." }
        }
        $UpdateProgress.Invoke("Restore Complete", 100, "Successfully restored from: $BackupRootPath"); Write-Host "--- LOCAL Restore Operation Finished ---"; return $true
    } catch { $errorMessage = "LOCAL Restore Failed: $($_.Exception.Message)"; Write-Error $errorMessage; $UpdateProgress.Invoke("Restore Failed", -1, $errorMessage); return $false }
}


# --- Express Mode Function ---
function Execute-ExpressModeLogic {
    param(
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$Credential
    )

    Write-Host "Executing Express Mode Logic..."
    $targetDevice = $null
    $mappedDriveLetter = "X" # Or choose another available letter
    $localTempTransferDir = Join-Path $env:TEMP "ExpressTransfer_$(Get-Date -Format 'yyyyMMddHHmmss')"
    $localBackupBaseDir = "C:\LocalData" # Where the final backup copy will reside
    $transferSuccess = $true # Assume success initially
    # FIX: Initialize $transferLog locally within the function
    $transferLog = [System.Collections.Generic.List[string]]::new()
    $transferLog.Add("Timestamp,Action,Status,Details")
    $logFilePath = $null # Initialize log file path variable
    $tempLogPath = $null # Initialize temp log file path variable

    # FIX: Use local $transferLog variable
    Function LogTransfer ($Action, $Status, $Details) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $safeDetails = $Details -replace '"', '""'
        $logEntry = """$timestamp"",""$Action"",""$Status"",""$safeDetails"""
        $transferLog.Add($logEntry) # Use local variable
        Write-Host "LOG: $Action - $Status - $Details"
    }

    try {
        # --- 1. Get Target Device ---
        $targetDevice = Read-Host "Enter the IP address or Hostname of the target remote device"
        if ([string]::IsNullOrWhiteSpace($targetDevice)) {
            throw "Target device cannot be empty."
        }
        LogTransfer "Input" "Info" "Target device specified: $targetDevice"

        # --- 2. Map Network Drive ---
        $uncPath = "\\$targetDevice\c$"
        Write-Host "Attempting to map $uncPath to $mappedDriveLetter`: " -NoNewline
        LogTransfer "Map Drive" "Attempt" "Mapping $uncPath to $mappedDriveLetter`:"
        if (Test-Path "${mappedDriveLetter}:") { Write-Warning "Drive $mappedDriveLetter`: exists. Removing..."; Remove-PSDrive -Name $mappedDriveLetter -Force -EA SilentlyContinue }
        New-PSDrive -Name $mappedDriveLetter -PSProvider FileSystem -Root $uncPath -Credential $Credential -ErrorAction Stop | Out-Null
        Write-Host "Successfully mapped $uncPath to $mappedDriveLetter`:" -ForegroundColor Green
        LogTransfer "Map Drive" "Success" "Successfully mapped $uncPath to $mappedDriveLetter`:"

        # --- 3. Get Remote Logged-on Username ---
        Write-Host "Attempting to identify logged-on user on '$targetDevice'..."
        LogTransfer "Get Remote User" "Attempt" "Querying Win32_ComputerSystem on $targetDevice"
        $remoteUsername = $null
        try {
            $remoteResult = Invoke-Command -ComputerName $targetDevice -Credential $Credential -ScriptBlock { (Get-CimInstance -ClassName Win32_ComputerSystem).UserName } -ErrorAction Stop
            if ($remoteResult -and -not [string]::IsNullOrWhiteSpace($remoteResult)) { $remoteUsername = ($remoteResult -split '\\')[-1]; Write-Host "Identified remote user: '$remoteUsername'" -ForegroundColor Green; LogTransfer "Get Remote User" "Success" "Identified remote user: $remoteUsername" }
            else { throw "Could not retrieve username from Win32_ComputerSystem or it was empty." }
        } catch {
            Write-Warning "Failed auto-detect remote user: $($_.Exception.Message)"; LogTransfer "Get Remote User" "Warning" "Failed WMI/Invoke: $($_.Exception.Message)"
            $remoteUsername = Read-Host "Could not auto-detect remote user. Please enter the Windows username logged into '$targetDevice'"; if ([string]::IsNullOrWhiteSpace($remoteUsername)) { throw "Remote username required." }; LogTransfer "Get Remote User" "Manual Input" "User provided: $remoteUsername"
        }

        # --- 4. Define Remote Paths and Local Destinations ---
        $remoteUserProfile = "$mappedDriveLetter`:\Users\$remoteUsername"; Write-Host "Remote user profile path: $remoteUserProfile"; LogTransfer "Path Setup" "Info" "Remote profile path: $remoteUserProfile"
        if (-not (Test-Path -LiteralPath $remoteUserProfile)) { throw "Remote user profile path '$remoteUserProfile' not found." }
        Write-Host "Identifying remote files/folders for user '$remoteUsername'..."; LogTransfer "Gather Paths" "Attempt" "Resolving templates for $remoteUserProfile"
        $remotePathsToTransfer = Resolve-UserPathTemplates -UserProfilePath $remoteUserProfile -Templates $script:pathTemplates; LogTransfer "Gather Paths" "Info" "Found $($remotePathsToTransfer.Count) template items."
        if ($remotePathsToTransfer.Count -eq 0) { Write-Warning "No template files/folders found for '$remoteUsername' on '$targetDevice'."; LogTransfer "Gather Paths" "Warning" "No template paths resolved." }
        Write-Host "Creating local temp dir: $localTempTransferDir"; LogTransfer "Setup" "Info" "Creating local temp dir: $localTempTransferDir"; New-Item -Path $localTempTransferDir -ItemType Directory -Force -EA Stop | Out-Null

        # --- 5. Transfer Files/Folders ---
        Write-Host "Starting file/folder transfer from '$targetDevice' to '$localTempTransferDir'..."
        $totalItems = $remotePathsToTransfer.Count; $currentItem = 0; $errorsDuringTransfer = $false
        if ($totalItems -gt 0) {
            foreach ($item in $remotePathsToTransfer) {
                $currentItem++; $progress = [int](($currentItem / $totalItems) * 50) + 10; Write-Progress -Activity "Transferring Files from $targetDevice" -Status "[$progress%]: Copying $($item.Name)" -PercentComplete $progress
                $sourcePath = $item.Path; $destRelativePath = $item.SourceKey -replace '[^a-zA-Z0-9_]', '_'; if ($item.Type -eq 'File') { $destRelativePath = Join-Path $destRelativePath (Split-Path $sourcePath -Leaf) }; $destinationPath = Join-Path $localTempTransferDir $destRelativePath
                Write-Host "  Copying '$($item.Name)' from '$sourcePath' to '$destinationPath'"; LogTransfer "File Transfer" "Attempt" "Copying $($item.SourceKey) from $sourcePath"
                try { $destDir = Split-Path $destinationPath -Parent; if (-not (Test-Path $destDir)) { New-Item -Path $destDir -ItemType Directory -Force -EA Stop | Out-Null }; Copy-Item -Path $sourcePath -Destination $destinationPath -Recurse:($item.Type -eq 'Folder') -Force -EA Stop; LogTransfer "File Transfer" "Success" "Copied $($item.SourceKey) to $destinationPath" }
                catch { Write-Warning "  Failed copy '$($item.Name)': $($_.Exception.Message)"; LogTransfer "File Transfer" "Error" "Failed copy $($item.SourceKey): $($_.Exception.Message)"; $errorsDuringTransfer = $true }
            }
            Write-Progress -Activity "Transferring Files from $targetDevice" -Completed
            if ($errorsDuringTransfer) { Write-Warning "Errors occurred during file/folder transfer." } else { Write-Host "File/folder transfer completed." -ForegroundColor Green }
        } else { Write-Host "Skipping file/folder transfer - no items."; Write-Progress -Activity "Transferring Files from $targetDevice" -Status "[60%]: No files" -PercentComplete 60 }

        # --- 6. Transfer Settings (Network Drives / Printers) ---
        Write-Progress -Activity "Transferring Settings" -Status "[60%]: Drives..." -PercentComplete 60; Write-Host "Getting mapped drives from $targetDevice"; LogTransfer "Get Drives" "Attempt" "Querying remote drives"
        $remoteDrivesCsvPath = Join-Path $localTempTransferDir "Drives.csv"
        try { Invoke-Command -ComputerName $targetDevice -Credential $Credential -ScriptBlock { Get-WmiObject -Class Win32_MappedLogicalDisk -EA Stop | Select-Object Name, ProviderName } -EA Stop | Export-Csv -Path $remoteDrivesCsvPath -NoTypeInformation -Encoding UTF8 -EA Stop; Write-Host "Saved remote drives." -ForegroundColor Green; LogTransfer "Get Drives" "Success" "Saved to $remoteDrivesCsvPath" }
        catch { Write-Warning "Failed get drives from ${targetDevice}: $($_.Exception.Message)"; LogTransfer "Get Drives" "Error" "Failed: $($_.Exception.Message)" }
        Write-Progress -Activity "Transferring Settings" -Status "[75%]: Printers..." -PercentComplete 75; Write-Host "Getting network printers from $targetDevice"; LogTransfer "Get Printers" "Attempt" "Querying remote printers"
        $remotePrintersTxtPath = Join-Path $localTempTransferDir "Printers.txt"
        try { Invoke-Command -ComputerName $targetDevice -Credential $Credential -ScriptBlock { Get-WmiObject -Class Win32_Printer -Filter "Local=False AND Network=True" -EA Stop | Select-Object -ExpandProperty Name } -EA Stop | Set-Content -Path $remotePrintersTxtPath -Encoding UTF8 -EA Stop; Write-Host "Saved remote printers." -ForegroundColor Green; LogTransfer "Get Printers" "Success" "Saved to $remotePrintersTxtPath" }
        catch { Write-Warning "Failed get printers from ${targetDevice}: $($_.Exception.Message)"; LogTransfer "Get Printers" "Error" "Failed: $($_.Exception.Message)" }
        Write-Progress -Activity "Transferring Settings" -Completed

        # --- 7. Post-Transfer Remote Actions (GPUpdate, ConfigMgr) ---
        Write-Host "Executing post-transfer actions remotely on '$targetDevice'..."; LogTransfer "Remote Actions" "Attempt" "Running GPUpdate/ConfigMgr on $targetDevice"; Write-Progress -Activity "Remote Actions" -Status "[85%]: GPUpdate/ConfigMgr..." -PercentComplete 85
        try {
            Invoke-Command -ComputerName $targetDevice -Credential $Credential -ScriptBlock {
                function Set-GPupdate { Write-Host "REMOTE: GPUpdate..."; try { $p = Start-Process cmd.exe '/c gpupdate /force' -PassThru -Wait -EA Stop; if ($p.ExitCode -eq 0) { Write-Host "REMOTE: GPUpdate OK." -ForegroundColor Green } else { Write-Warning "REMOTE: GPUpdate ExitCode $($p.ExitCode)." } } catch { Write-Error "REMOTE: GPUpdate failed: $($_.Exception.Message)" } }
                function Start-ConfigManagerActions { Write-Host "REMOTE: ConfigMgr Actions..."; $ccmExec="C:\Windows\CCM\ccmexec.exe"; $ns="root\ccm\clientsdk"; $cl="CCM_ClientUtilities"; $m="TriggerSchedule"; $ok=$false; $cimOk=$false; $acts=@(@{ID='{...21}';N='MachPol'},@{ID='{...22}';N='UserPol'},@{ID='{...01}';N='HInv'},@{ID='{...02}';N='SInv'},@{ID='{...113}';N='SUScan'},@{ID='{...101}';N='HInvCol'},@{ID='{...108}';N='SUAssEval'},@{ID='{...102}';N='SInvCol'}); $svc=Get-Service CcmExec -EA SilentlyContinue; if(!$svc){Write-Warning "REMOTE: CcmExec Svc N/F"; return $false}; if($svc.Status -ne 'Running'){Write-Warning "REMOTE: CcmExec Svc not running"; return $false}; Write-Host "REMOTE: CcmExec Svc OK."; Write-Host "REMOTE: Try CIM..."; $cimMethOk=$true; try { if(Get-CimClass -N $ns -C $cl -EA SilentlyContinue){ foreach($a in $acts){ Write-Host "REMOTE: . CIM $($a.N)"; try{Invoke-CimMethod -N $ns -C $cl -M $m -Args @{sScheduleID=$a.ID} -EA Stop}catch{Write-Warning "REMOTE: . CIM $($a.N) failed: $($_.Exception.Message)";$cimMethOk=$false} } if($cimMethOk){$cimOk=$true;$ok=$true;Write-Host "REMOTE: CIM OK." -ForegroundColor Green}else{Write-Warning "REMOTE: CIM failed some."} }else{Write-Warning "REMOTE: CIM class N/F.";$cimMethOk=$false} }catch{Write-Error "REMOTE: CIM error: $($_.Exception.Message)";$cimMethOk=$false}; if(!$cimOk){ Write-Host "REMOTE: Try ccmexec.exe..."; if(Test-Path $ccmExec){ $exeOk=$true; foreach($a in $acts){ Write-Host "REMOTE: . EXE $($a.N)"; try{$p=Start-Process $ccmExec "-TriggerSchedule $($a.ID)" -NoNewWindow -PassThru -Wait -EA Stop;if($p.ExitCode -ne 0){Write-Warning "REMOTE: . EXE $($a.N) ExitCode $($p.ExitCode)."}}catch{Write-Warning "REMOTE: . EXE $($a.N) failed: $($_.Exception.Message)";$exeOk=$false} } if($exeOk){$ok=$true;Write-Host "REMOTE: EXE OK." -ForegroundColor Green}else{Write-Warning "REMOTE: EXE failed some."} }else{Write-Warning "REMOTE: ccmexec.exe N/F."} }; if($ok){Write-Host "REMOTE: ConfigMgr Actions OK." -ForegroundColor Green}else{Write-Warning "REMOTE: ConfigMgr Actions Failed."}; return $ok }
                Set-GPupdate; Start-ConfigManagerActions; Write-Host "REMOTE: Actions finished."
            } -ErrorAction Stop; Write-Host "Executed remote actions." -ForegroundColor Green; LogTransfer "Remote Actions" "Success" "Executed GPUpdate/ConfigMgr remotely."
        } catch { Write-Warning "Failed remote actions on ${targetDevice}: $($_.Exception.Message)"; LogTransfer "Remote Actions" "Error" "Failed: $($_.Exception.Message)" }
        Write-Progress -Activity "Remote Actions" -Completed

        # --- 8. Create Final Local Backup Copy ---
        Write-Progress -Activity "Creating Local Backup" -Status "[90%]: Copying..." -PercentComplete 90; $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"; $finalBackupDir = Join-Path $localBackupBaseDir "TransferBackup_${remoteUsername}_from_($targetDevice -replace '[^a-zA-Z0-9_.-]','_')_$timestamp"; Write-Host "Creating local backup copy in '$finalBackupDir'..."; LogTransfer "Local Backup" "Attempt" "Copying from $localTempTransferDir to $finalBackupDir"
        if ((Get-ChildItem -Path $localTempTransferDir -ErrorAction SilentlyContinue).Count -gt 0) {
            try { Copy-Item -Path $localTempTransferDir -Destination $finalBackupDir -Recurse -Force -EA Stop; Write-Host "Local backup copy created." -ForegroundColor Green; LogTransfer "Local Backup" "Success" "Created backup at $finalBackupDir" }
            catch { Write-Warning "Failed create final backup '$finalBackupDir': $($_.Exception.Message)"; LogTransfer "Local Backup" "Error" "Failed: $($_.Exception.Message)"; $transferSuccess = $false }
        } else { Write-Warning "No items transferred to temp dir for backup."; LogTransfer "Local Backup" "Warning" "Skipped - No items in $localTempTransferDir" }
        Write-Progress -Activity "Creating Local Backup" -Completed

        # --- 9. Final Log Saving ---
        if (-not ([string]::IsNullOrEmpty($finalBackupDir)) -and -not (Test-Path $finalBackupDir)) { New-Item -Path $finalBackupDir -ItemType Directory -Force -EA SilentlyContinue | Out-Null }
        if (-not ([string]::IsNullOrEmpty($finalBackupDir)) -and (Test-Path $finalBackupDir)) { $logFilePath = Join-Path $finalBackupDir "TransferLog.csv" }
        try { if (-not ([string]::IsNullOrEmpty($logFilePath))) { $transferLog | Set-Content -Path $logFilePath -Encoding UTF8 -Force; Write-Host "Log saved: $logFilePath" } else { $tempLogPath = Join-Path $env:TEMP "ExpressTransfer_Log_$(Get-Date -Format 'yyyyMMddHHmmss').csv"; $transferLog | Set-Content -Path $tempLogPath -Encoding UTF8 -Force; Write-Warning "Saved log to temp: $tempLogPath"; LogTransfer "Save Log" "Warning" "Saved log to temp: $tempLogPath" } }
        catch { Write-Error "CRITICAL: Failed save log: $($_.Exception.Message)"; LogTransfer "Save Log" "Fatal Error" "Failed save log: $($_.Exception.Message)" }
        Write-Progress -Activity "Express Transfer Complete" -Status "[100%]: Finished." -PercentComplete 100; Write-Host "--- REMOTE Transfer Operation Finished ---"

    } catch { $errorMessage = "Express Transfer Failed: $($_.Exception.Message)"; Write-Error $errorMessage; LogTransfer "Overall Status" "Fatal Error" $errorMessage; $transferSuccess = $false }
    finally {
        # --- 10. Cleanup ---
        if (Test-Path "${mappedDriveLetter}:") { Write-Host "Removing mapped drive $mappedDriveLetter`:" -NoNewline; try { Remove-PSDrive -Name $mappedDriveLetter -Force -EA Stop; Write-Host " OK." -ForegroundColor Green; LogTransfer "Cleanup" "Success" "Removed drive $mappedDriveLetter`:" } catch { Write-Warning " FAILED remove drive $mappedDriveLetter`: $($_.Exception.Message)"; LogTransfer "Cleanup" "Error" "Failed remove drive $mappedDriveLetter`: $($_.Exception.Message)" } }
        if (Test-Path $localTempTransferDir) { Write-Host "Removing temp dir: $localTempTransferDir"; Remove-Item -Path $localTempTransferDir -Recurse -Force -EA SilentlyContinue; LogTransfer "Cleanup" "Info" "Removed temp dir $localTempTransferDir" }
        $finalLogPathMessage = "Log saving failed"; if (-not ([string]::IsNullOrEmpty($logFilePath)) -and (Test-Path $logFilePath)) { $finalLogPathMessage = $logFilePath } elseif (-not ([string]::IsNullOrEmpty($tempLogPath)) -and (Test-Path $tempLogPath)) { $finalLogPathMessage = $tempLogPath }
        if ($transferSuccess) { Write-Host "Express Transfer process completed." -ForegroundColor Green; [System.Windows.MessageBox]::Show("Express Transfer completed. Check log: `n$finalLogPathMessage", "Express Transfer Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) }
        else { Write-Error "Express Transfer failed. Check log: $finalLogPathMessage"; [System.Windows.MessageBox]::Show("Express Transfer failed. Check log: `n$finalLogPathMessage", "Express Transfer Failed", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error); throw "Express Transfer failed. Check log: $finalLogPathMessage" }
    }
}

#endregion Functions

# --- Main Execution ---
Write-Host "--- Script Starting ---"
Clear-Variable -Name updateJob -Scope Script -ErrorAction SilentlyContinue
$operationMode = 'Cancel' # Default

# Ensure Default Path Exists
if (-not (Test-Path $script:DefaultPath)) { Write-Host "Default path '$($script:DefaultPath)' not found. Creating."; try { New-Item -Path $script:DefaultPath -ItemType Directory -Force -EA Stop | Out-Null; Write-Host "Default path created." } catch { Write-Warning "Could not create default path: $($script:DefaultPath)." } }

try {
    Write-Host "Calling Show-ModeDialog..."
    $operationMode = Show-ModeDialog

    switch ($operationMode) {
        'Backup' { Write-Host "Mode: Backup"; Show-MainWindow -Mode 'Backup' }
        'Restore' { Write-Host "Mode: Restore"; Show-MainWindow -Mode 'Restore' } # Background job started in Show-ModeDialog
        'Express' {
            Write-Host "Mode: Express. Checking elevation..."
            $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
            if (-not $isAdmin) {
                Write-Warning "Elevation required. Relaunching..."
                try { Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs -EA Stop }
                catch { Write-Error "Failed relaunch: $($_.Exception.Message). Run as Admin manually."; [System.Windows.MessageBox]::Show("Failed relaunch. Run as Admin manually.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) }
                Exit
            } else {
                Write-Host "Already Admin. Getting credentials..."
                $credential = $null; try { $credential = Get-Credential -UserName "$env:USERDOMAIN\$env:USERNAME" -Message "Enter ADMIN credentials for REMOTE machine" } catch { throw "Failed get credentials: $($_.Exception.Message)" }
                if ($credential -eq $null) { throw "Credentials required for Express mode." }
                Execute-ExpressModeLogic -Credential $credential
            }
        }
        'Cancel' { Write-Host "Operation Cancelled." }
        Default { Write-Error "Invalid mode: $operationMode" }
    }
} catch {
    $errorMessage = "FATAL ERROR: $($_.Exception.Message)"
    if ($_.FullyQualifiedErrorId -match 'Express Transfer failed') { Write-Host "Express mode failed. See previous messages." -ForegroundColor Red }
    else { Write-Error $errorMessage -EA Continue; try { [System.Windows.MessageBox]::Show($errorMessage, "Fatal Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) } catch {} }
} finally {
    if ($operationMode -eq 'Restore' -and (Get-Variable -Name updateJob -Scope Script -EA SilentlyContinue) -ne $null -and $script:updateJob -ne $null) {
        Write-Host "`n--- Waiting for background Restore updates job... ---" -ForegroundColor Yellow; Wait-Job $script:updateJob | Out-Null; Write-Host "--- Background Job Output: ---" -ForegroundColor Yellow; Receive-Job $script:updateJob; Remove-Job $script:updateJob; Write-Host "--- End Background Job Output ---" -ForegroundColor Yellow
    } elseif ($operationMode -ne 'Express' -and $operationMode -ne 'Cancel') { Write-Host "`nNo background update job started/needed." -ForegroundColor Gray }
}

Write-Host "--- Script Execution Finished ---"
if ($Host.Name -eq 'ConsoleHost' -and -not $psISE -and $env:TERM_PROGRAM -ne 'vscode') { Write-Host "Press Enter to exit..." -ForegroundColor Yellow; Read-Host }